using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.VisualStudio.VCProjectEngine;

namespace VSESwitchEnv
{
    public class EnvVar(string name, string value)
    {
        public string Name { get; set; } = name.Split(':')[0].Trim();
        public string Value { get; set; } = value.Trim();
        public bool Define { get; set; } = name.Contains(":");

        public static implicit operator bool(EnvVar i) => !string.IsNullOrEmpty(i.Name);
    }

    internal class EnvProject(string name, VCProject project = null)
    {
        public readonly string Name = name;
        public readonly Dictionary<string, List<EnvVar>> Data = [];
        public string Selected { get; private set; } = null;
        public VCProject Project { get; set; } = project;

        public static implicit operator bool(EnvProject i) => i.Data.Count > 0;

        public string UpdateProps(string selected)
        {
            if (selected == null || !Data.ContainsKey(selected))
                return null;

            Selected = selected;
            var subProps = CreatePropsFile();
            if (Project == null)
                return subProps;

            ApplyProps(subProps);
            return null;
        }

        public void ApplyProps(string subProps)
        {
            if (Project == null)
                return;

            var ext = subProps.Substring(subProps.Length - 8);
            var isShared = ext[1] == 's';
            var mainProps = Path.Combine(EnvSolution.ExtDir, $"{Name}.props");

            bool modified = false;
            foreach (VCConfiguration vcConf in Project.Configurations as IVCCollection)
            {
                VCPropertySheet mainSheet = null;
                foreach (VCPropertySheet sheet in vcConf.PropertySheets as IVCCollection)
                {
                    if (sheet.PropertySheetFile == mainProps)
                    {
                        mainSheet = sheet;
                        break;
                    }
                }
                if (mainSheet != null)
                {
                    if (!modified)
                    {
                        foreach (VCPropertySheet sheet in mainSheet.PropertySheets as IVCCollection)
                        {
                            if (sheet.PropertySheetFile.EndsWith(ext))
                                mainSheet.RemovePropertySheet(sheet);
                        }
                        var added = mainSheet.AddPropertySheet(subProps);
                        if (isShared)
                            try { mainSheet.MovePropertySheet(added, false); } catch { }
                        mainSheet.Save();
                        modified = true;
                    }
                    vcConf.AddPropertySheet(mainProps);
                }
            }
            Project.Save();
        }

        public void PushData(string key, EnvVar var)
        {
            if (!Data.TryGetValue(key, out var data))
                Data[key] = data = [];

            data.Add(var);
        }

        private string CreatePropsFile()
        {
            XNamespace ns = Consts.PropXmlNs;
            var userMacros = new XElement(ns + "PropertyGroup", new XAttribute("Label", "UserMacros"));
            var buildMacros = new XElement(ns + "ItemGroup");
            string prepDefs = string.Empty;

            if (Data.TryGetValue(Selected, out var data))
            {
                foreach (var v in data)
                {
                    var value = v.Value;
                    if (value.StartsWith("\"") && value.EndsWith("\""))
                        value = value.Trim('"');
                    else if (value.StartsWith("R\"(") && value.EndsWith(")\""))
                        value = value.Substring(3, value.Length - 5);
                    userMacros.Add(new XElement(ns + v.Name, value));
                    buildMacros.Add(new XElement(ns + "BuildMacro", new XAttribute("Include", v.Name),
                        new XElement(ns + "Value", $"$({v.Name})"),
                        new XElement(ns + "EnvironmentVariable", "true")
                    ));
                    if (v.Define)
                    {
                        string defName = Regex.Replace(v.Name, "(?<=[a-z0-9])([A-Z])", "_$1");
                        defName = Regex.Replace(defName, "([A-Z])([A-Z][a-z])", "$1_$2").ToUpperInvariant();
                        prepDefs += defName + $"={v.Value};";
                    }
                }
            }

            var root = new XElement(ns + "Project", new XAttribute("ToolsVersion", "Current"), userMacros, buildMacros);
            if (prepDefs != string.Empty)
            {
                root.Add(new XElement(ns + "ItemDefinitionGroup", new XAttribute("Condition", "'$(VCProjectVersion)' != ''"),
                    new XElement(ns + "ClCompile", new XElement(ns + "PreprocessorDefinitions", prepDefs + "%(PreprocessorDefinitions)"))));
            }
            XDocument doc = new(root);

            var ext = $"{(Project != null ? "x" : "s")}.props";
            foreach (var f in Directory.GetFiles(EnvSolution.ExtDir, $"{Name}.*.{ext}"))
                try { File.Delete(f); } catch { }

            var props = Path.Combine(EnvSolution.ExtDir, $"{Name}.{DateTime.Now:HHmmss}.{ext}");
            doc.Save(props);
            return props;
        }

        public static void FixProjects(string solFile)
        {
            foreach (var line in File.ReadAllLines(solFile))
            {
                if (line.Trim().StartsWith("Project("))
                {
                    var parts = line.Split(',');
                    if (parts.Length >= 2)
                    {
                        var relPath = parts[1].Trim().Trim('"');
                        if (relPath.EndsWith(".vcxproj", StringComparison.OrdinalIgnoreCase))
                        {
                            if (EnvSolution.Projects.ContainsKey(Consts.SharedName) || EnvSolution.Projects.ContainsKey(Path.GetFileNameWithoutExtension(relPath)))
                                FixPropsImport(Path.Combine(EnvSolution.SolDir, relPath));
                        }
                    }
                }
            }
        }

        private static void FixPropsImport(string projFile)
        {
            bool isDirty = false;
            var doc = XDocument.Load(projFile);
            var ns = doc.Root.GetDefaultNamespace();

            var props = $"{Path.GetFileNameWithoutExtension(projFile)}.props";
            var propRelPath = $"{Consts.ExtRelPath}\\{props}";
            var propAttr = new XAttribute("Project", $"$(SolutionDir)\\{propRelPath}");
            var propImport = new XElement(ns + "Import", propAttr, new XAttribute("Condition", $"exists('{propAttr.Value}')"));

            var sheetNodes = doc.Descendants(ns + "ImportGroup").Where(item => (string)item.Attribute("Label") == "PropertySheets");
            foreach (var sheetNode in sheetNodes)
            {
                var propNode = sheetNode.Descendants(ns + "Import").Where(item => item.Attribute("Project").Value.EndsWith(propRelPath)).FirstOrDefault();
                if (propNode == null)
                {
                    sheetNode.Add(propImport);
                    isDirty = true;
                }
            }
            if (isDirty)
                doc.Save(projFile);

            ns = Consts.PropXmlNs;
            doc = new XDocument(new XElement(ns + "Project", new XAttribute("ToolsVersion", "Current"), new XElement(ns + "ImportGroup", new XAttribute("Label", "PropertySheets"))));
            doc.Save(Path.Combine(EnvSolution.ExtDir, props));
        }
    }
}