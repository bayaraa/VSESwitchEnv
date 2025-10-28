using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.VCProjectEngine;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace VSESwitchEnv
{
    internal class EnvProject(string name) : EnvData
    {
        public readonly string Name = name;
        private VCProject _project = null;

        public bool IsReady() => _project != null && !IsEmpty();
        public void SetProject(VCProject project) => _project = project;

        public void UpdateProps()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!Services.DTE.Solution.IsOpen || _project == null)
                return;

            bool isDirty = false;
            var mainProps = Path.Combine(EnvSolution.ExtDir, $"{Name}.props");
            foreach (VCConfiguration vcConf in _project.Configurations as IVCCollection)
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
                    if (!isDirty)
                    {
                        foreach (VCPropertySheet sheet in mainSheet.PropertySheets as IVCCollection)
                            mainSheet.RemovePropertySheet(sheet);
                        mainSheet.AddPropertySheet(CreateProps());
                        mainSheet.Save();
                        isDirty = true;
                    }
                    vcConf.AddPropertySheet(mainProps);
                }
            }
            if (isDirty)
                _project.Save();
        }

        private string CreateProps()
        {
            XNamespace ns = Consts.PropXmlNs;
            var userMacros = new XElement(ns + "PropertyGroup", new XAttribute("Label", "UserMacros"));
            var buildMacros = new XElement(ns + "ItemGroup");
            string prepDefs = string.Empty;

            foreach (var v in EnvSolution.SharedData.Current().Concat(Current()))
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

            var root = new XElement(ns + "Project", new XAttribute("ToolsVersion", "Current"), userMacros, buildMacros);
            if (prepDefs != string.Empty)
            {
                root.Add(new XElement(ns + "ItemDefinitionGroup", new XAttribute("Condition", "'$(VCProjectVersion)' != ''"),
                    new XElement(ns + "ClCompile", new XElement(ns + "PreprocessorDefinitions", prepDefs + "%(PreprocessorDefinitions)"))));
            }
            XDocument doc = new(root);

            DeleteProps();
            var props = Path.Combine(EnvSolution.ExtDir, $"{Name}.{DateTime.Now:HHmmss}.props");
            doc.Save(props);
            return props;
        }

        private void DeleteProps()
        {
            foreach (var f in Directory.GetFiles(EnvSolution.ExtDir, $"{Name}.*.props"))
                try { File.Delete(f); } catch { }
        }

        public void FixPropsImport(string projFile)
        {
            bool isDirty = false;
            var doc = XDocument.Load(projFile);
            var ns = doc.Root.GetDefaultNamespace();
            var propRelPath = $"{Consts.ExtRelPath}\\{Name}.props";
            var propAttr = new XAttribute("Project", $"$(SolutionDir)\\{propRelPath}");
            var propImport = new XElement(ns + "Import", propAttr, new XAttribute("Condition", $"exists('{propAttr.Value}')"));

            var sheetNodes = doc.Descendants(ns + "ImportGroup").Where(item => (string?)item.Attribute("Label") == "PropertySheets");
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
        }

        public void CreateMainProps()
        {
            XNamespace ns = Consts.PropXmlNs;
            var doc = new XDocument(new XElement(ns + "Project", new XAttribute("ToolsVersion", "Current"),
                new XElement(ns + "ImportGroup", new XAttribute("Label", "PropertySheets"))));
            doc.Save(Path.Combine(EnvSolution.ExtDir, $"{Name}.props"));
        }

        public void DeleteMainProps()
        {
            DeleteProps();
            try { File.Delete(Path.Combine(EnvSolution.ExtDir, $"{Name}.props")); } catch { }
        }
    }
}