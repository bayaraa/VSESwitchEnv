using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.VCProjectEngine;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace VSESwitchEnv
{
    internal sealed class EnvSwitcher
    {
        public static readonly Guid CommandSet = new("39032d76-f68d-41eb-9051-60cb49430b48");

        public class Variable
        {
            public string Name { get; set; } = null;
            public string Value { get; set; } = null;
            public bool Define { get; set; } = false;
        }

        public class EnvConfig
        {
            public string Selected { get; set; } = null;
            public Dictionary<string, List<Variable>> Options { get; set; } = [];

            public static implicit operator bool(EnvConfig i) => i.Options.Count > 0;
        }

        private readonly DTE2 _dte;
        private readonly string _extName = "VSESwitchEnv";
        private readonly OleMenuCommand _menuCommand;
        private OutputWindowPane _outputPane = null;
        private bool _outputPaneActive = false;
        private string _extDir = null;
        private string _stateJson = null;

        private readonly EnvConfig _defaultConfig = new();
        private readonly Dictionary<string, EnvConfig> _envConfigs = [];
        private string _currentProjName = null;

        public EnvSwitcher(DTE2 dte, OleMenuCommandService commandService)
        {
            _dte = dte;
            _menuCommand = new OleMenuCommand(new EventHandler(OnMenuSelect), new CommandID(CommandSet, 0x101));
            commandService.AddCommand(_menuCommand);
            commandService.AddCommand(new OleMenuCommand(new EventHandler(OnMenuGetList), new CommandID(CommandSet, 0x102)));
        }

        public void OnBeforeOpenProject(object sender, Microsoft.VisualStudio.Shell.Events.BeforeOpenProjectEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (string.IsNullOrEmpty(_dte.Solution?.FullName) || string.IsNullOrEmpty(e?.Filename))
                return;

            if (File.Exists(e.Filename) && e.Filename.EndsWith(".vcxproj"))
            {
                bool isDirty = false;
                var doc = XDocument.Load(e.Filename);
                var ns = doc.Root.GetDefaultNamespace();

                var propName = $"{Path.GetFileNameWithoutExtension(e.Filename)}.props";
                var propRelPath = $".vs\\{_extName}\\{propName}";
                var propPathAttr = new XAttribute("Project", $"$(SolutionDir)\\{propRelPath}");
                var propImport = new XElement(ns + "Import", propPathAttr, new XAttribute("Condition", $"exists('{propPathAttr.Value}')"));

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
                    doc.Save(e.Filename);

                var solDir = new DirectoryInfo(Path.GetDirectoryName(_dte.Solution.FullName)!);
                if (solDir == null)
                    return;

                _extDir = Path.Combine(solDir.FullName, $".vs\\{_extName}\\");
                Directory.CreateDirectory(_extDir);

                ns = "http://schemas.microsoft.com/developer/msbuild/2003";
                doc = new XDocument(new XElement(ns + "Project", new XAttribute("ToolsVersion", "Current"), new XElement(ns + "ImportGroup", new XAttribute("Label", "PropertySheets"))));
                doc.Save(Path.Combine(_extDir, propName));
            }
        }

        public void OnSolutionOpen(object sender = null, EventArgs e = null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (string.IsNullOrEmpty(_dte.Solution?.FullName) || !_dte.Solution.IsOpen)
                return;

            var solDir = new DirectoryInfo(Path.GetDirectoryName(_dte.Solution.FullName)!);
            if (solDir == null)
                return;

            var configFile = solDir.GetFiles(".vseswitchenv", SearchOption.TopDirectoryOnly).FirstOrDefault();
            if (configFile == null)
            {
                configFile = solDir.GetFiles(".editorconfig", SearchOption.TopDirectoryOnly).FirstOrDefault();
                if (configFile == null)
                    return;
            }
            var configStr = File.ReadAllText(configFile.FullName);
            if (string.IsNullOrEmpty(configStr))
                return;

            foreach (Project project in _dte.Solution.Projects)
            {
                if (TryAsVCProject(project, out var vcProj, out var pName))
                    _envConfigs[pName] = new EnvConfig();
            }

            Action<string, string, Variable> appendVariable = (projName, envName, variable) =>
            {
                if (_envConfigs.ContainsKey(projName))
                {
                    if (!_envConfigs[projName].Options.ContainsKey(envName))
                        _envConfigs[projName].Options.Add(envName, new List<Variable>());
                    _envConfigs[projName].Options[envName].Add(variable);
                }
            };

            using (StringReader reader = new StringReader(configStr))
            {
                bool isEnv = false;
                string line = null;
                string envName = null;
                string[] projectNames = [];

                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (line == string.Empty)
                        continue;

                    if (line.StartsWith("["))
                    {
                        if (line.StartsWith("[env:") && line.EndsWith("]"))
                        {
                            line = line.Replace("[env:", "").Replace("]", "");
                            var segments = line.Split(new char[] { '|' }, 2);
                            envName = segments[0].Trim();
                            projectNames = segments.Length == 2 ? segments[1].Trim().Split(',') : [];
                            isEnv = true;
                            continue;
                        }
                        isEnv = false;
                    }
                    if (isEnv && line != string.Empty && line[0] != '#')
                    {
                        var segments = line.Split(new char[] { '=' }, 2);
                        if (segments.Length == 2)
                        {
                            var variable = new Variable();
                            variable.Name = segments[0].Trim();
                            variable.Value = segments[1].Trim();
                            if (variable.Name.Contains(":"))
                            {
                                variable.Define = true;
                                variable.Name = variable.Name.Split(':')[0].Trim();
                                if (variable.Name == string.Empty)
                                {
                                    Output("Invalid format at: " + line);
                                    continue;
                                }
                            }

                            if (projectNames.Length > 0)
                            {
                                foreach (var projName in projectNames)
                                    appendVariable(projName, envName, variable);
                            }
                            else
                            {
                                foreach (var cfg in _envConfigs)
                                    appendVariable(cfg.Key, envName, variable);
                            }
                        }
                    }
                }
            }

            _extDir = Path.Combine(solDir.FullName, $".vs\\{_extName}\\");
            Directory.CreateDirectory(_extDir);

            _stateJson = Path.Combine(_extDir, "state.json");
            if (File.Exists(_stateJson))
            {
                var data = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(_stateJson));
                foreach (var v in data!)
                {
                    if (_envConfigs.ContainsKey(v.Key) && v.Value != null && _envConfigs[v.Key].Options.ContainsKey(v.Value))
                        _envConfigs[v.Key].Selected = v.Value;
                }
            }
            foreach (var envConf in _envConfigs)
            {
                UpdateProjectProps(envConf.Key, envConf.Value);
#if DEBUG
                Output("Project: " + envConf.Key + " | selected: " + envConf.Value.Selected);
                foreach (var env in envConf.Value.Options)
                {
                    Output("  Env name: " + env.Key);
                    foreach (var v in env.Value)
                        Output("    var: " + v.Define + " | " + v.Name + " = " + v.Value);
                }
#endif
            }

            EnableMenu();
        }

        public void OnSolutionClose(object sender = null, EventArgs e = null)
        {
            DisableMenu();
            _envConfigs.Clear();
            _currentProjName = null;
        }

        public void OnWindowCreated(Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SetProject(window.Project);
        }

        public void OnWindowActivated(Window gotFocus, Window lostFocus)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SetProject(gotFocus.Project);
        }

        public void OnBuildBegin(vsBuildScope Scope, vsBuildAction Action) => DisableMenu();
        public void OnBuildDone(vsBuildScope Scope, vsBuildAction Action)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            EnableMenu();
        }

        public void OnEnterRunMode(dbgEventReason reason) => DisableMenu();
        public void OnEnterDesignMode(dbgEventReason reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            EnableMenu();
        }

        private void OnMenuSelect(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (e is OleMenuCmdEventArgs eventArgs)
            {
                string selected = eventArgs.InValue as string;
                IntPtr vOut = eventArgs.OutValue;

                var config = GetConfig();
                if (config)
                {
                    if (vOut != IntPtr.Zero)
                        Marshal.GetNativeVariantForObject(config.Selected, vOut);
                    else if (selected != null)
                    {
                        config.Selected = selected;
                        UpdateProjectProps(_currentProjName, config);
                    }
                }
            }
        }

        private void OnMenuGetList(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (e is OleMenuCmdEventArgs eventArgs)
            {
                object inParam = eventArgs.InValue;
                IntPtr vOut = eventArgs.OutValue;

                if (inParam == null && vOut != IntPtr.Zero)
                {
                    var options = new List<string>();
                    foreach (var opt in GetConfig().Options)
                        options.Add(opt.Key);
                    Marshal.GetNativeVariantForObject(options.ToArray(), vOut);
                }
            }
        }

        private void UpdateProjectProps(string projectName, EnvConfig config)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!FindVCProjectByName(projectName, out var vcProj))
                return;

            var propName = $"{projectName}.props";
            foreach (var f in Directory.GetFiles(_extDir, $"{projectName}.*.props"))
            {
                if (Path.GetFileName(f) != propName)
                    try { File.Delete(f); } catch { }
            }
            var propsFile = Path.Combine(_extDir, $"{propName}");
            var subPropsFile = Path.Combine(_extDir, $"{projectName}.{DateTime.Now.ToString("HHmmss")}.props");

            XNamespace ns = "http://schemas.microsoft.com/developer/msbuild/2003";
            var userMacros = new XElement(ns + "PropertyGroup", new XAttribute("Label", "UserMacros"));
            var buildMacros = new XElement(ns + "ItemGroup");
            string prepDefStr = string.Empty;

            if (config.Selected != null && config.Options.ContainsKey(config.Selected))
            {
                foreach (var v in config.Options[config.Selected])
                {
                    userMacros.Add(new XElement(ns + v.Name, v.Value));
                    buildMacros.Add(new XElement(ns + "BuildMacro", new XAttribute("Include", v.Name),
                        new XElement(ns + "Value", $"$({v.Name})"),
                        new XElement(ns + "EnvironmentVariable", "true")
                    ));
                    if (v.Define)
                    {
                        string defName = Regex.Replace(v.Name, "(?<=[a-z0-9])([A-Z])", "_$1");
                        defName = Regex.Replace(defName, "([A-Z])([A-Z][a-z])", "$1_$2").ToUpperInvariant();
                        prepDefStr += defName + $"=\"$({v.Name})\";";
                    }
                }
            }
            var root = new XElement(ns + "Project", new XAttribute("ToolsVersion", "Current"), userMacros, buildMacros);
            if (prepDefStr != string.Empty)
                root.Add(new XElement(ns + "ItemDefinitionGroup", new XAttribute("Condition", "'$(VCProjectVersion)' != ''"),
                    new XElement(ns + "ClCompile", new XElement(ns + "PreprocessorDefinitions", prepDefStr + "%(PreprocessorDefinitions)"))));

            XDocument doc = new XDocument(root);
            doc.Save(subPropsFile);

            foreach (VCConfiguration vcConf in vcProj.Configurations as IVCCollection)
            {
                foreach (VCPropertySheet sheet in vcConf.PropertySheets as IVCCollection)
                {
                    if (sheet.PropertySheetFile == propsFile)
                    {
                        foreach (VCPropertySheet subSheet in sheet.PropertySheets as IVCCollection)
                            sheet.RemovePropertySheet(subSheet);
                        sheet.AddPropertySheet(subPropsFile);
                        sheet.Save();

                        vcConf.AddPropertySheet(propsFile);
                        break;
                    }
                }
            }
            vcProj.Save();

            var states = new Dictionary<string, string>();
            foreach (var envConf in _envConfigs)
                states.Add(envConf.Key, envConf.Value.Selected);
            var options = new JsonSerializerOptions { WriteIndented = true };

            try
            {
                File.WriteAllText(_stateJson, JsonSerializer.Serialize(states, options));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception during write state.json: {ex}");
            }

            if (!string.IsNullOrEmpty(config.Selected))
                Output($"{projectName}: Environment \"{config.Selected}\" selected.");
        }

        private void SetProject(Project project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (TryAsVCProject(project, out var vcProj, out var pName) && _currentProjName != pName)
            {
                DisableMenu();
                _currentProjName = pName;
                EnableMenu();
            }
        }

        private void DisableMenu() => _menuCommand.Enabled = false;
        private void EnableMenu()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _menuCommand.Enabled = GetConfig();
        }

        private EnvConfig GetConfig()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_currentProjName == null)
            {
                foreach (Project project in _dte.Solution.Projects)
                {
                    if (TryAsVCProject(project, out var vcProj, out var pName))
                    {
                        _currentProjName = pName;
                        break;
                    }
                }
            }
            if (_currentProjName != null && _envConfigs.ContainsKey(_currentProjName))
                return _envConfigs[_currentProjName];

            return _defaultConfig;
        }

        private bool TryAsVCProject(Project project, out VCProject vcProj, out string name)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            vcProj = null;
            name = null;
            if (project == null)
                return false;

            try
            {
                vcProj = project.Object as VCProject;
                if (vcProj != null && !string.IsNullOrEmpty(project.FileName))
                {
                    name = Path.GetFileNameWithoutExtension(project.FileName);
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool FindVCProjectByName(string name, out VCProject vcProj)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (Project project in _dte.Solution.Projects)
            {
                if (TryAsVCProject(project, out vcProj, out var pName) && pName == name)
                    return true;
            }
            vcProj = null;
            return false;
        }

        private void Output(string line, bool clear = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CreateOutputPane();

            if (clear)
                ClearOutput();

            if (!string.IsNullOrEmpty(line))
            {
                if (!_outputPaneActive)
                {
                    _outputPane.Activate();
                    _outputPaneActive = true;
                }
                _outputPane.OutputString(line + Environment.NewLine);
            }

        }

        private void ClearOutput()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CreateOutputPane();
            _outputPane.Clear();
        }

        private void CreateOutputPane()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_outputPane == null)
            {
                if (_dte?.ToolWindows?.OutputWindow == null)
                    return;

                try { _outputPane = _dte.ToolWindows.OutputWindow.OutputWindowPanes.Item(_extName); }
                catch { _outputPane = _dte.ToolWindows.OutputWindow.OutputWindowPanes.Add(_extName); }
            }
        }
    }
}
