using EnvDTE;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace VSESwitchEnv
{
    internal class MenuCommand
    {
        private readonly OleMenuCommand _cmd;
        private bool _wasEnabled = false;
        private readonly int _id;

        public MenuCommand(int id)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _id = id;

            var guid = new Guid(Consts.CmdSetGuid);
            Services.Command.AddCommand(new OleMenuCommand(new EventHandler(OnMenuGetList), new CommandID(guid, _id)));

            _cmd = new OleMenuCommand(new EventHandler(OnMenuSelect), new CommandID(guid, _id + 1));
            Services.Command.AddCommand(_cmd);


        Services.DTE.Events.BuildEvents.OnBuildBegin += OnBuildBegin;
            Services.DTE.Events.BuildEvents.OnBuildDone += OnBuildDone;
            Services.DTE.Events.DebuggerEvents.OnEnterRunMode += OnEnterRunMode;
            Services.DTE.Events.DebuggerEvents.OnEnterDesignMode += OnEnterDesignMode;

            Disable();
        }

        public void Enable() => _wasEnabled = _cmd.Enabled = true;
        public void Disable() => _wasEnabled = _cmd.Enabled = false;

        private void OnMenuGetList(object sender, EventArgs e)
        {
            if (e is OleMenuCmdEventArgs eventArgs)
            {
                object inParam = eventArgs.InValue;
                IntPtr vOut = eventArgs.OutValue;

                if (inParam == null && vOut != IntPtr.Zero)
                {
                    var options = new List<string>();
                    options.Add("[Shared]");
                    options.Add("➔ Option1");
                    options.Add("　 Option2");
                    options.Add("[Project]");
                    options.Add("");
                    options.Add("➔ Option1");
                    options.Add("　 Option2");
                    options.Add("﹡ Option2");
                    options.Add("ーーーー");
                    options.Add("　 Option2");
                    options.Add("✔️ Option3");
                    options.Add("🗸 Option4");

                    if (EnvSolution.GetProject(out var project, _id == Consts.CmdIDShared))
                    {
                        foreach (var env in project.Data)
                            options.Add(env.Key);
                    }
                    Marshal.GetNativeVariantForObject(options.ToArray(), vOut);
                }
            }
        }

        private void OnMenuSelect(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (e is OleMenuCmdEventArgs eventArgs)
            {
                string selected = eventArgs.InValue as string;
                IntPtr vOut = eventArgs.OutValue;

                if (vOut != IntPtr.Zero)
                {
                    if (EnvSolution.GetProject(out var project, _id == Consts.CmdIDShared))
                        Marshal.GetNativeVariantForObject(project.Selected, vOut);
                }
                else if (selected != null)
                    EnvSolution.OnMenuSelect(_id, selected);
            }
        }

        private void OnBuildBegin(vsBuildScope Scope, vsBuildAction Action) => _cmd.Enabled = false;
        private void OnEnterRunMode(dbgEventReason reason) => _cmd.Enabled = false;

        private void OnBuildDone(vsBuildScope Scope, vsBuildAction Action) => _cmd.Enabled = _wasEnabled;
        private void OnEnterDesignMode(dbgEventReason reason) => _cmd.Enabled = _wasEnabled;
    }
}
