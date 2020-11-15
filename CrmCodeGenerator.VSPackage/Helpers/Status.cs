using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace CrmCodeGenerator.VSPackage.Helpers
{
    public static class Status
    {
        public static void Update(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            //Configuration.Instance.DTE.ExecuteCommand("View.Output");
            var dte = Package.GetGlobalService(typeof(SDTE)) as EnvDTE.DTE;
            var win = dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
            win.Visible = true;

            var outputWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
            if (outputWindow == null)
                return;

            Guid guidGeneral = Microsoft.VisualStudio.VSConstants.OutputWindowPaneGuid.GeneralPane_guid;
            IVsOutputWindowPane pane;
            outputWindow.CreatePane(guidGeneral, "Crm Code Generator", 1, 0);
            outputWindow.GetPane(guidGeneral, out pane);
            pane.Activate();
            pane.OutputString(message);
            pane.OutputString("\n");
            pane.FlushToTaskList();
            System.Windows.Forms.Application.DoEvents();
        }

        public static void Clear()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var outputWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
            if (outputWindow == null)
                return;

            Guid guidGeneral = Microsoft.VisualStudio.VSConstants.OutputWindowPaneGuid.GeneralPane_guid;
            IVsOutputWindowPane pane;
            outputWindow.CreatePane(guidGeneral, "Crm Code Generator", 1, 0);
            outputWindow.GetPane(guidGeneral, out pane);
            pane.Clear();
            pane.FlushToTaskList();
            System.Windows.Forms.Application.DoEvents();
        }
    }
}
