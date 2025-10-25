using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace VSESwitchEnv
{
    [Guid("d6b8c8d1-fff8-4b8c-9f8c-ed6eda389220")]
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideMenuResource("Menus.ctmenu", 1)]

    public sealed class VSESwitchEnvPackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            var commandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;

            OutputPane.Initialize(dte, Consts.ExtName);
            EnvSolution.Initialize(dte, commandService);
        }
    }
    internal static class Consts
    {
        public const string ExtName = "VSESwitchEnv";
        public const string ExtRelPath = $".vs\\{ExtName}";
        public const string SharedName = "Shared";
        public const string PropXmlNs = "http://schemas.microsoft.com/developer/msbuild/2003";
    }
}