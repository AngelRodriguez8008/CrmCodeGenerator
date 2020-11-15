using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using CrmCodeGenerator.VSPackage.Helpers;
using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace CrmCodeGenerator.VSPackage
{
    internal sealed class ApplyCustomTool
    {
        private readonly AsyncPackage _package;
        private IServiceProvider ServiceProvider => _package;

        private readonly DTE _dte;
        private XrmCodeGenerator _generator;
        
        public static ApplyCustomTool Instance
        {
            get;
            private set;
        }

        private ApplyCustomTool(AsyncPackage package,IMenuCommandService commandService, DTE dte)
        {
            _dte = dte;
            _package = package;

            var cmdId = new CommandID(GuidList.guidCrmCodeGenerator_VSPackageCmdSet, (int)PkgCmdIDList.ApplyCustomToolCmdId);
            var command = new OleMenuCommand(Execute, cmdId)
            {
                // This will defer visibility control to the VisibilityConstraints section in the .vsct file
                Supported = false
            };
            commandService.AddCommand(command);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            if (Instance != null)
                return;
            
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            Assumes.Present(package);

            var dte = await package.GetServiceAsync(typeof(DTE)) as DTE;
            Assumes.Present(dte);

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as IMenuCommandService;
            Assumes.Present(commandService);

            Instance = new ApplyCustomTool(package, commandService, dte);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var item = _dte.SelectedItems.Item(1).ProjectItem;
            if (item == null)
                return;

            var filePath = item.Document?.FullName ?? item.FileNames[1];
            if (string.IsNullOrWhiteSpace(filePath))
                return;
            
            //item.Properties.Item("CustomTool").Value = XrmCodeGenerator.Name;
            
            _generator = _generator ?? new XrmCodeGenerator();
            var generateCode = _generator.GenerateCode(filePath);
            if (string.IsNullOrWhiteSpace(generateCode))
                return;

            var extention = _generator.GetDefaultExtension();
            var outputFileName = Path.ChangeExtension(filePath, extention);
          
            byte[] bytes = Encoding.UTF8.GetBytes(generateCode);
            File.WriteAllBytes(outputFileName, bytes);
        }

        //public void Execute(object sender, EventArgs e)
        //{
        //    ThreadHelper.ThrowIfNotOnUIThread();

        //    string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", typeof(ApplyCustomTool).FullName);
          
        //    IntPtr hierarchyPointer, selectionContainerPointer;
        //    object selectedObject = null;
        //    IVsMultiItemSelect multiItemSelect;
        //    uint projectItemId;

        //    IVsMonitorSelection monitorSelection =
        //        (IVsMonitorSelection)Package.GetGlobalService(
        //            typeof(SVsShellMonitorSelection));

        //    monitorSelection.GetCurrentSelection(out hierarchyPointer,
        //        out projectItemId,
        //        out multiItemSelect,
        //        out selectionContainerPointer);

        //    var selectedHierarchy = Marshal.GetTypedObjectForIUnknown(hierarchyPointer, typeof(IVsHierarchy)) as IVsHierarchy;

        //    if (selectedHierarchy != null)
        //    {
        //        ErrorHandler.ThrowOnFailure(selectedHierarchy.GetProperty(
        //            projectItemId,
        //            (int)__VSHPROPID.VSHPROPID_ExtObject,
        //            out selectedObject));
        //    }

        //    var selectedItem = selectedObject as ProjectItem;
        //    var filePath  =  selectedItem?.FileNames[1];

        //    // Show a message box to prove we were here
        //    VsShellUtilities.ShowMessageBox(
        //        ServiceProvider,
        //        message,
        //        filePath,
        //        OLEMSGICON.OLEMSGICON_INFO,
        //        OLEMSGBUTTON.OLEMSGBUTTON_OK,
        //        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        //}
    }
}
