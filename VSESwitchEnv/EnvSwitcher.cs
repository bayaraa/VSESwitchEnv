using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.VCProjectEngine;

namespace VSESwitchEnv
{
    internal sealed class EnvSwitcher
    {
        public static readonly Guid CommandSet = new Guid("39032d76-f68d-41eb-9051-60cb49430b48");

        private readonly DTE2 dte;
        private readonly OleMenuCommand menuCommand;
        private readonly string configFileName = ".vseswitchenv";
        private readonly string paneName = "VSESwitchEnv";

        private OutputWindowPane outputPane = null;
        private string selectedOption = String.Empty;

        public EnvSwitcher(DTE2 dte, OleMenuCommandService commandService)
        {
            this.dte = dte;

            menuCommand = new OleMenuCommand(new EventHandler(OnSelect), new CommandID(CommandSet, 0x101));
            commandService.AddCommand(menuCommand);
            commandService.AddCommand(new OleMenuCommand(new EventHandler(OnGetList), new CommandID(CommandSet, 0x102)));
            menuCommand.Enabled = false;
        }

        private void SwitchEnv(string selected)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (selectedOption == selected || selected == String.Empty)
                return;

            var config = GetConfig();
            if (config == String.Empty)
                return;

            selectedOption = selected;
            var userMacros = new Dictionary<string, string>();
            using (StringReader reader = new StringReader(config))
            {
                string line;
                bool found = false;
                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (line == String.Empty)
                        continue;

                    if (!found && line == "[" + selectedOption + "]")
                    {
                        found = true;
                        continue;
                    }
                    if (found)
                    {
                        if (line.StartsWith("["))
                            break;

                        var segments = line.Split(new char[] { '=' }, 2);
                        if (segments.Length == 2)
                            userMacros.Add(segments[0].Trim(), segments[1].Trim());
                    }
                }
            }

            var propertySheet = GetPropertySheet();
            if (propertySheet == null)
                return;

            propertySheet.RemoveAllUserMacros();
            foreach (var userMacro in userMacros)
            {
                var macro = propertySheet.AddUserMacro(userMacro.Key, userMacro.Value);
                macro.PerformEnvironmentSet = true;
            }
            propertySheet.AddUserMacro("SelectedEnv", selectedOption.ToString());

            if (propertySheet.IsDirty)
                propertySheet.Save();
        }

        private void OnSelect(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (e is OleMenuCmdEventArgs eventArgs)
            {
                string selected = eventArgs.InValue as string;
                IntPtr vOut = eventArgs.OutValue;

                if (vOut != IntPtr.Zero)
                    Marshal.GetNativeVariantForObject(selectedOption, vOut);
                else if (selected != null)
                    SwitchEnv(selected);
            }
        }

        private void OnGetList(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (e is OleMenuCmdEventArgs eventArgs)
            {
                object inParam = eventArgs.InValue;
                IntPtr vOut = eventArgs.OutValue;

                if (inParam == null && vOut != IntPtr.Zero)
                {
                    var options = new List<string>();
                    var config = GetConfig();
                    if (config  != String.Empty)
                    {
                        using (StringReader reader = new StringReader(config))
                        {
                            string line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                if (line.StartsWith("[") && line.EndsWith("]"))
                                    options.Add(line.Trim(new char[] { '[', ']' }));
                            }
                        }
                    }
                    Marshal.GetNativeVariantForObject(options.ToArray(), vOut);
                }
            }
        }

        public void OnOpen(object sender = null, EventArgs e = null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            menuCommand.Enabled = GetConfig() != String.Empty;
            if (menuCommand.Enabled)
            {
                var propertySheet = GetPropertySheet(true);
                if (propertySheet == null)
                    return;

                foreach (VCUserMacro macro in propertySheet.UserMacros)
                {
                    if (macro.Name == "SelectedEnv") {
                        selectedOption = macro.Value;
                        break;
                    }
                }
            }
        }

        public void OnClose(object sender = null, EventArgs e = null)
        {
            menuCommand.Enabled = false;
        }

        private string GetConfig()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!dte.Solution.IsOpen || String.IsNullOrEmpty(dte.Solution.FullName))
                return String.Empty;

            var filePath = Path.Combine(Path.GetDirectoryName(dte.Solution.FullName), configFileName);
            if (!File.Exists(filePath))
                return String.Empty;

            return System.IO.File.ReadAllText(filePath);
        }

        private VCPropertySheet GetPropertySheet(bool readOnly = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!dte.Solution.IsOpen || String.IsNullOrEmpty(dte.Solution.FullName))
                return null;

            var solutionName = Path.GetFileNameWithoutExtension(dte.Solution.FullName);
            var filePath = Path.Combine(Path.GetDirectoryName(dte.Solution.FullName), solutionName + ".props");

            if (!readOnly && !File.Exists(filePath))
            {
                var content = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Project ToolsVersion=\"Current\" xmlns=\"http://schemas.microsoft.com/developer/msbuild/2003\">" +
                    "<ImportGroup Label=\"PropertySheets\" />" +
                    "</Project>";
                System.IO.File.WriteAllText(filePath, content);
            }

            VCPropertySheet propertySheet = null;
            foreach (Project project in dte.Solution.Projects)
            {
                if (!(project.Object is VCProject vcProject))
                    continue;

                foreach (VCConfiguration configuration in vcProject.Configurations as IVCCollection)
                {
                    if (readOnly)
                    {
                        foreach (VCPropertySheet sheet in configuration.PropertySheets as IVCCollection)
                        {
                            if (!sheet.IsSystemPropertySheet && sheet.PropertySheetFile == filePath)
                                return sheet;
                        }
                    }
                    else
                        propertySheet = configuration.AddPropertySheet(filePath);
                }
            }
            return propertySheet;
        }

        private void Output(string line)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (line != String.Empty)
            {
                CreateOutputPane();
                outputPane.Activate();
                outputPane.OutputString(line + Environment.NewLine);
            }
        }

        private void ClearOutput()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            CreateOutputPane();
            outputPane.Clear();
        }

        private void CreateOutputPane()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (outputPane == null)
                outputPane = dte.ToolWindows.OutputWindow.OutputWindowPanes.Add(paneName);
        }
    }
}
