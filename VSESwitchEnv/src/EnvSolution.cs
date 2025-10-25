using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
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
        public static readonly Dictionary<string, EnvProject> Projects = new(StringComparer.OrdinalIgnoreCase);

        private static DTE2 _dte = null;
        private static string _curProject = null;
        private static bool _isEnabled = false;
        private static bool _wasEnabled = false;

        private static WindowEvents _windowEvents;
        private static BuildEvents _buildEvents;
        private static DebuggerEvents _debuggerEvents;

        public static void Initialize(DTE2 dte, OleMenuCommandService command)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _dte = dte;

            Events.SolutionEvents.OnBeforeOpenSolution += OnBeforeOpenSolution;
            Events.SolutionEvents.OnAfterOpenSolution += OnAfterOpenSolution;
            Events.SolutionEvents.OnBeforeCloseSolution += OnBeforeCloseSolution;
            _windowEvents = _dte.Events.WindowEvents;
            _windowEvents.WindowCreated += OnWindowCreated;
            _windowEvents.WindowActivated += OnWindowActivated;

            var guid = new Guid("39032d76-f68d-41eb-9051-60cb49430b48");
            command.AddCommand(new OleMenuCommand(new EventHandler(OnMenuGetList), new CommandID(guid, 0x101)));
            var cmd = new OleMenuCommand(new EventHandler(OnMenuSelect), new CommandID(guid, 0x102));
            command.AddCommand(cmd);
            cmd.BeforeQueryStatus += OnBeforeQueryStatus;

            _buildEvents = _dte.Events.BuildEvents;
            _buildEvents.OnBuildBegin += OnBuildBegin;
            _buildEvents.OnBuildDone += OnBuildDone;
            _debuggerEvents = _dte.Events.DebuggerEvents;
            _debuggerEvents.OnEnterRunMode += OnEnterRunMode;
            _debuggerEvents.OnEnterDesignMode += OnEnterDesignMode;
        }

        private static void OnBeforeOpenSolution(object sender, Events.BeforeOpenSolutionEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SolDir = Path.GetDirectoryName(e.SolutionFilename);
            ExtDir = Path.Combine(SolDir, Consts.ExtRelPath);
            Directory.CreateDirectory(ExtDir);

            if (Config.Load())
                EnvProject.FixProjects(e.SolutionFilename);
        }

        private static void OnAfterOpenSolution(object sender, Events.OpenSolutionEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!_dte.Solution.IsOpen)
                return;

            foreach (Project proj in EnumerateProjects())
            {
                if (!string.IsNullOrEmpty(proj.FileName) && proj.Object is VCProject vcProj)
                {
                    var name = Path.GetFileNameWithoutExtension(proj.FileName);
                    if (Projects.TryGetValue(name, out var project))
                        project.Project = vcProj;
                    else
                        Projects.Add(name, new(name, vcProj));
                    _curProject ??= name;
                }
            }
            if (_curProject == null)
                return;

            Config.LoadStates();
            if (GetProject(out var p, true) || GetProject(out p))
                _isEnabled = true;
        }

        private static void OnBeforeCloseSolution(object sender, EventArgs e)
        {
            Projects.Clear();
            _curProject = null;
            _isEnabled = false;
        }

        private static void OnWindowCreated(Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SetProject(window);
        }

        private static void OnWindowActivated(Window gotFocus, Window lostFocus)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SetProject(gotFocus);
        }

        private static void OnMenuGetList(object sender, EventArgs e)
        {
            if (e is not OleMenuCmdEventArgs eventArgs)
                return;

            if (eventArgs.InValue == null && eventArgs.OutValue != IntPtr.Zero)
            {
                var options = new List<string>();
                if (GetProject(out var shared, true))
                {
                    foreach (var env in shared.Data)
                        options.Add($"{(shared.Selected == env.Key ? "⏵" : "　")}{env.Key} ");
                }
                if (GetProject(out var project))
                {
                    if (options.Count > 0)
                        options.Add("⋯⋯⋯⋯⋯⋯");
                    foreach (var env in project.Data)
                        options.Add($"{(project.Selected == env.Key ? "⏵" : "　")}{env.Key}");
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
                if (GetProject(out var shared, true))
                    text += shared.Selected;
                if (GetProject(out var project))
                {
                    if (text != string.Empty)
                        text += " / ";
                    text += project.Selected;
                }
                text += text == string.Empty ? " ---" : "  ";
                Marshal.GetNativeVariantForObject(text, eventArgs.OutValue);
                return;
            }

            string selected = eventArgs.InValue as string;
            if (selected != null && !selected.StartsWith("⋯"))
            {
                if (GetProject(out var project, selected.EndsWith(" ")))
                {
                    selected = selected.Remove(0, 1).Trim();
                    if (selected != project.Selected)
                    {
                        UpdateProps(project, selected);
                        OutputPane.Write($"[{project.Name}] environment updated: {project.Selected}.");
                        Config.SaveStates();
                    }
                }
            }
        }

        private static void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand cmd = sender as OleMenuCommand;
            if (cmd != null)
                cmd.Enabled = _isEnabled;
        }

        private static void OnBuildBegin(vsBuildScope Scope, vsBuildAction Action)
        {
            _wasEnabled = _isEnabled;
            _isEnabled = false;
        }
        private static void OnEnterRunMode(dbgEventReason reason)
        {
            _wasEnabled = _isEnabled;
            _isEnabled = false;
        }

        private static void OnBuildDone(vsBuildScope Scope, vsBuildAction Action) => _isEnabled = _wasEnabled;
        private static void OnEnterDesignMode(dbgEventReason reason) => _isEnabled = _wasEnabled;

        public static void UpdateProps(EnvProject project, string selected)
        {
            var props = project.UpdateProps(selected);
            if (project.Project == null && props != null)
            {
                foreach (var p in Projects)
                    p.Value.ApplyProps(props);
            }
        }

        public static bool GetProject(out EnvProject project, bool shared = false)
        {
            var name = shared ? Consts.SharedName : _curProject;
            if (name != null && Projects.TryGetValue(name, out project))
                return project;
            project = null;
            return false;
        }

        private static void SetProject(Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Project project = window?.Project ?? _dte.ActiveDocument?.ProjectItem?.ContainingProject;
            if (string.IsNullOrEmpty(project?.FileName))
                return;

            var name = Path.GetFileNameWithoutExtension(project.FileName);
            if (_curProject == name)
                return;

            _curProject = null;
            _isEnabled = false;
            if (Projects.TryGetValue(name, out var p) && p)
            {
                _curProject = name;
                _isEnabled = true;
            }
            else if (GetProject(out p, true) && p)
                _isEnabled = true;
        }

        private static IEnumerable<Project> EnumerateProjects()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (Project p in _dte.Solution.Projects)
            {
                foreach (var sub in EnumerateProjectsRecursive(p))
                    yield return sub;
            }
        }

        private static IEnumerable<Project> EnumerateProjectsRecursive(Project parent)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parent == null)
                yield break;

            if (parent.Kind == Constants.vsProjectKindSolutionItems || parent.Kind == ProjectKinds.vsProjectKindSolutionFolder)
            {
                if (parent.ProjectItems != null)
                {
                    foreach (ProjectItem item in parent.ProjectItems)
                    {
                        var sub = item.SubProject;
                        if (sub != null)
                        {
                            foreach (var p in EnumerateProjectsRecursive(sub))
                                yield return p;
                        }
                    }
                }
                yield break;
            }
            yield return parent;
        }
    }
}