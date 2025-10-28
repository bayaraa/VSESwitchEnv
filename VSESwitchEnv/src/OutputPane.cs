using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;

namespace VSESwitchEnv
{
    internal static class OutputPane
    {
        private static OutputWindowPane _pane = null;
        private static string _name = null;
        private static bool _isActive = false;

        public static void Initialize(string name)
        {
            _name = name;
        }

        public static void Write(string text, bool clear = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            EnsurePane();

            if (clear)
                _pane.Clear();

            if (string.IsNullOrEmpty(text))
                return;

            if (!_isActive)
            {
                _pane.Activate();
                _isActive = true;
            }

            _pane.OutputString(text + Environment.NewLine);
        }

        private static void EnsurePane()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_pane != null)
                return;

            if (Services.DTE?.ToolWindows?.OutputWindow == null)
                return;

            try
            {
                _pane = Services.DTE.ToolWindows.OutputWindow.OutputWindowPanes.Item(_name);
            }
            catch
            {
                _pane = Services.DTE.ToolWindows.OutputWindow.OutputWindowPanes.Add(_name);
            }
        }
    }
}