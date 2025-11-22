using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.VCProjectEngine;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Runtime.InteropServices;
using Events = Microsoft.VisualStudio.Shell.Events;

namespace VSESwitchEnv
{
    internal static class EnvSolution
    {
        public static string SolDir { get; private set; } = null;
        public static string ExtDir { get; private set; } = null;

        public static readonly EnvData SharedData = new();
        public static readonly Dictionary<string, EnvProject> Projects = new(StringComparer.OrdinalIgnoreCase);

        private static string _curProject = null;
        private static bool _disabled = false;

        private static WindowEvents _windowEvents;
        private static BuildEvents _buildEvents;
        private static DebuggerEvents _debuggerEvents;

        public static void Initialize()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Events.SolutionEvents.OnBeforeOpenSolution += OnBeforeOpenSolution;
            Events.SolutionEvents.OnBeforeOpenProject += OnBeforeOpenProject;
            Events.SolutionEvents.OnAfterOpenProject += OnAfterOpenProject;
            Events.SolutionEvents.OnAfterOpenSolution += OnAfterOpenSolution;
            Events.SolutionEvents.OnBeforeCloseSolution += OnBeforeCloseSolution;
            Events.SolutionEvents.OnBeforeCloseProject += OnBeforeCloseProject;

            _windowEvents = Services.DTE.Events.WindowEvents;
            _windowEvents.WindowCreated += (window) => SetProject(window);
            _windowEvents.WindowActivated += (window, _) => SetProject(window);

            var guid = new Guid("39032d76-f68d-41eb-9051-60cb49430b48");
            Services.Cmd.AddCommand(new OleMenuCommand(new EventHandler(OnMenuGetList), new CommandID(guid, 0x101)));
            var cmd = new OleMenuCommand(new EventHandler(OnMenuSelect), new CommandID(guid, 0x102));
            cmd.BeforeQueryStatus += (s, _) => ((OleMenuCommand)s).Enabled = (_curProject != null || !SharedData.IsEmpty()) && !_disabled;
            Services.Cmd.AddCommand(cmd);

            _buildEvents = Services.DTE.Events.BuildEvents;
            _buildEvents.OnBuildBegin += (_, _) => _disabled = true;
            _buildEvents.OnBuildDone += (_, _) => _disabled = false;
            _debuggerEvents = Services.DTE.Events.DebuggerEvents;
            _debuggerEvents.OnEnterRunMode += (_) => _disabled = true;
            _debuggerEvents.OnEnterDesignMode += (_) => _disabled = false;
        }

        private static void OnBeforeOpenSolution(object sender, Events.BeforeOpenSolutionEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SolDir = Path.GetDirectoryName(e.SolutionFilename);
            ExtDir = Path.Combine(SolDir, Consts.ExtRelPath);
            Directory.CreateDirectory(ExtDir);
            EnvConfig.Load();
        }

        private static void OnBeforeOpenProject(object sender, Events.BeforeOpenProjectEventArgs e)
        {
            if (e.Filename == null || !e.Filename.EndsWith(".vcxproj", StringComparison.OrdinalIgnoreCase))
                return;

            var name = Path.GetFileNameWithoutExtension(e.Filename);
            if (!Projects.ContainsKey(name) && !SharedData.IsEmpty())
                Projects.Add(name, new(name));
            if (Projects.TryGetValue(name, out var project))
            {
                project.FixPropsImport(e.Filename);
                project.CreateMainProps();
            }
        }

        private static void OnAfterOpenProject(object sender, Events.OpenProjectEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (GetHierarchyProject(e.Hierarchy, out var proj) && proj.Object is VCProject vcProj)
            {
                var name = Path.GetFileNameWithoutExtension(proj.FileName);
                if (Projects.TryGetValue(name, out var project))
                {
                    project.SetProject(vcProj);
                    project.UpdateProps();
                }
            }
        }

        private static void OnAfterOpenSolution(object sender, Events.OpenSolutionEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (var p in Projects)
                p.Value.UpdateProps();
        }

        private static void OnBeforeCloseSolution(object sender, EventArgs e)
        {
            foreach (var p in Projects)
                p.Value.DeleteMainProps();
            Projects.Clear();

            SharedData.Clear();
            _curProject = null;
        }

        private static void OnBeforeCloseProject(object sender, Events.CloseProjectEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (GetHierarchyProject(e.Hierarchy, out var proj))
            {
                var name = Path.GetFileNameWithoutExtension(proj.FileName);
                if (Projects.TryGetValue(name, out var project))
                {
                    project.SetProject(null);
                    project.DeleteMainProps();
                    if (_curProject == name)
                        _curProject = null;
                }
            }
        }

        private static void OnMenuGetList(object sender, EventArgs e)
        {
            if (e is not OleMenuCmdEventArgs eventArgs)
                return;

            if (eventArgs.InValue == null && eventArgs.OutValue != IntPtr.Zero)
            {
                var options = new List<string>();
                foreach (var env in SharedData.List())
                    options.Add($"{(SharedData.Selected() == env.Key ? "⏵" : "　")}{env.Key} ");
                if (GetProject(out var project))
                {
                    if (options.Count > 0)
                        options.Add("⋯⋯⋯⋯⋯⋯");
                    foreach (var env in project.List())
                        options.Add($"{(project.Selected() == env.Key ? "⏵" : "　")}{env.Key}");
                }
                Marshal.GetNativeVariantForObject(options.ToArray(), eventArgs.OutValue);
            }
        }

        private static void OnMenuSelect(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (e is not OleMenuCmdEventArgs eventArgs)
                return;

            if (eventArgs.OutValue != IntPtr.Zero)
            {
                string text = string.Empty;
                if (!SharedData.IsEmpty())
                    text += SharedData.Selected();
                if (GetProject(out var project))
                {
                    if (!string.IsNullOrEmpty(text))
                        text += " / ";
                    text += project.Selected();
                }
                text += text == string.Empty ? " ---" : "  ";
                Marshal.GetNativeVariantForObject(text, eventArgs.OutValue);
                return;
            }

            string selected = eventArgs.InValue as string;
            if (selected != null && !selected.StartsWith("⋯"))
            {
                selected = selected.Remove(0, 1);
                if (selected.EndsWith(" "))
                {
                    if (SharedData.Selected(selected.Trim()))
                    {
                        foreach (var p in Projects)
                            p.Value.UpdateProps();
                        OutputPane.Write($"[{Consts.SharedName}] environment updated: {SharedData.Selected()}.");
                    }

                }
                else if (GetProject(out var project))
                {
                    if (project.Selected(selected))
                    {
                        project.UpdateProps();
                        OutputPane.Write($"[{project.Name}] environment updated: {project.Selected()}.");
                    }
                }
                EnvConfig.Save();
            }
        }

        private static bool GetHierarchyProject(IVsHierarchy hierarchy, out Project project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (hierarchy != null)
            {
                var res = hierarchy.GetProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_ExtObject, out object extObj);
                if (ErrorHandler.Succeeded(res) && extObj is Project proj && !string.IsNullOrEmpty(proj.FileName) && proj.Object is VCProject)
                {
                    project = proj;
                    return true;
                }
            }
            project = null;
            return false;
        }

        private static bool GetProject(out EnvProject project)
        {
            if (_curProject != null && Projects.TryGetValue(_curProject, out project) && project.IsReady())
                return true;
            project = null;
            return false;
        }

        private static void SetProject(Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Project project = window?.Project ?? Services.DTE.ActiveDocument?.ProjectItem?.ContainingProject;
            if (string.IsNullOrEmpty(project?.FileName))
                return;

            var name = Path.GetFileNameWithoutExtension(project.FileName);
            if (_curProject == name)
                return;

            _curProject = null;
            if (Projects.TryGetValue(name, out var p) && p.IsReady())
                _curProject = name;
        }
    }
}