using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using SolutionEvents = Microsoft.VisualStudio.Shell.Events.SolutionEvents;
using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace VSESwitchEnv
{
    [Guid(PackageGuidString)]
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionOpening_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideMenuResource("Menus.ctmenu", 1)]

    public sealed class VSESwitchEnvPackage : AsyncPackage
    {
        public const string PackageGuidString = "d6b8c8d1-fff8-4b8c-9f8c-ed6eda389220";

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            Assumes.Present(dte);
            var commandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Assumes.Present(commandService);

            EnvSwitcher envSwitcher = new(dte, commandService);

            var solutionService = await GetServiceAsync(typeof(SVsSolution)) as IVsSolution;
            Assumes.Present(solutionService);
            ErrorHandler.ThrowOnFailure(solutionService.GetProperty((int)__VSPROPID.VSPROPID_IsSolutionOpen, out object isSolutionOpen));

            if ((bool)isSolutionOpen)
                envSwitcher.OnSolutionOpen();

            SolutionEvents.OnAfterOpenSolution += envSwitcher.OnSolutionOpen;
            SolutionEvents.OnAfterCloseSolution += envSwitcher.OnSolutionClose;
            SolutionEvents.OnBeforeOpenProject += envSwitcher.OnBeforeOpenProject;
            dte.Events.BuildEvents.OnBuildBegin += envSwitcher.OnBuildBegin;
            dte.Events.BuildEvents.OnBuildDone += envSwitcher.OnBuildDone;
            dte.Events.DebuggerEvents.OnEnterRunMode += envSwitcher.OnEnterRunMode;
            dte.Events.DebuggerEvents.OnEnterDesignMode += envSwitcher.OnEnterDesignMode;
            dte.Events.WindowEvents.WindowCreated += envSwitcher.OnWindowCreated;
            dte.Events.WindowEvents.WindowActivated += envSwitcher.OnWindowActivated;
        }
    }
}
