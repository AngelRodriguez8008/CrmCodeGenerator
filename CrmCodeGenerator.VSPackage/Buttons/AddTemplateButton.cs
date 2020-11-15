using System;
using System.ComponentModel.Design;
using System.IO;
using CrmCodeGenerator.VSPackage.Dialogs;
using CrmCodeGenerator.VSPackage.Helpers;
using CrmCodeGenerator.VSPackage.Model;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CrmCodeGenerator.VSPackage
{
    public sealed class AddTemplateButton
    {
        private readonly Settings settings = Configuration.Instance.Settings;
   
        private readonly AsyncPackage _package;
        private IServiceProvider ServiceProvider => _package;
        
        public static AddTemplateButton Instance
        {
            get;
            private set;
        }

        private AddTemplateButton(AsyncPackage package, IMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var cmdId = new CommandID(GuidList.guidCrmCodeGenerator_VSPackageCmdSet, (int)PkgCmdIDList.AddTemplateCmdId);
            var command = new OleMenuCommand(Execute, cmdId);

            commandService.AddCommand(command);
        }
        

        public static void Initialize(AsyncPackage package, IMenuCommandService commandService)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Instance = new AddTemplateButton(package, commandService);
        }
        
        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        public void Execute(object sender, EventArgs e)
        {
            try
            {
                AddTemplate();
                settings.IsActive = true;  // start saving the properties to the *.sln
            }
            catch (UserException uex)
            {
                VsShellUtilities.ShowMessageBox(ServiceProvider, uex.Message, "Error", OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
            catch (Exception ex)
            {
                var error = ex.Message + "\n" + ex.StackTrace;
                System.Windows.MessageBox.Show(error, "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
        
        private void AddTemplate()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = ServiceProvider.GetService(typeof(SDTE)) as DTE2;
            if (dte == null)
                return;

            var project = dte.GetSelectedProject();
            if (project == null || string.IsNullOrWhiteSpace(project.FullName))
            {
                throw new UserException("Please select a project first");
            }

            var m = new AddTemplate(dte, project);
            m.Closed += (sender, e) =>
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                // logic here Will be called after the child window is closed
                if (((AddTemplate)sender).Canceled == true)
                    return;

                var templatePath = Path.GetFullPath(Path.Combine(project.GetProjectDirectory(), m.Props.Template));  //GetFullpath removes un-needed relative paths  (ie if you are putting something in the solution directory)

                if (File.Exists(templatePath))
                {
                    var results = VsShellUtilities.ShowMessageBox(ServiceProvider, "'" + templatePath + "' already exists, are you sure you want to overwrite?", "Overwrite", OLEMSGICON.OLEMSGICON_QUERY, OLEMSGBUTTON.OLEMSGBUTTON_YESNOCANCEL, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    if (results != 6)
                        return;

                    //if the window is open we have to close it before we overwrite it.
                    var pi = project.GetProjectItem(m.Props.Template);
                    if(pi != null && pi.Document != null)
                        pi.Document.Close(vsSaveChanges.vsSaveChangesNo);
                }

                var templateSamplesPath = Path.Combine(DteHelper.AssemblyDirectory(), @"Resources\Templates");
                var defaultTemplatePath = Path.Combine(templateSamplesPath, m.DefaultTemplate.SelectedValue.ToString());
                if (!File.Exists(defaultTemplatePath))
                {
                    throw new UserException("T4Path: " + defaultTemplatePath + " is missing or you can access it.");
                }

                var dir = Path.GetDirectoryName(templatePath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                Status.Update("Adding " + templatePath + " to project");
                // When you add a TT file to visual studio, it will try to automatically compile it, 
                // if there is error (and there will be error because we have custom generator) 
                // the error will persit until you close Visual Studio. The solution is to add 
                // a blank file, then overwrite it
                // http://stackoverflow.com/questions/17993874/add-template-file-without-custom-tool-to-project-programmatically
                var blankTemplatePath = Path.Combine(DteHelper.AssemblyDirectory(), @"Resources\Templates\Blank.tt");
                File.Copy(blankTemplatePath, templatePath, true);

                var p = project.ProjectItems.AddFromFile(templatePath);
                p.Properties.SetValue("CustomTool", "");

                File.Copy(defaultTemplatePath, templatePath, true);
                p.Properties.SetValue("CustomTool", XrmCodeGenerator.Name);
            };
            m.ShowModal();
        }
    }
}