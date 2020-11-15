using EnvDTE;
using System;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;

namespace CrmCodeGenerator.VSPackage.Helpers
{
    public static class DteHelper
    {
        public static bool HasProjectItem(this Project project, string projectFile)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var projectItem = project.GetProjectItem(projectFile);
            return projectItem != null;
        }
        public static ProjectItem GetProjectItem(this Project project, string projectFile)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return GetProjectItemRecursive(project.ProjectItems, projectFile);
        }
        private static ProjectItem GetProjectItemRecursive(ProjectItems projectItems, string projectFile, string folder = "")
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            projectFile = projectFile.Replace(@"\", "/");

            foreach (ProjectItem projectItem in projectItems)
            {
                // if the name matches
                var current = projectItem.Name;
                if (current == projectFile || folder + current == projectFile)
                    return projectItem;

                var children = projectItem.ProjectItems;
                if (children != null && children.Count > 0)
                {
                    var result = GetProjectItemRecursive(children, projectFile, folder + current + "/");
                    if (result != null)
                        return result;
                }
            }
            return null;
        }
        public static string MakeRelative(string fromAbsolutePath, string toDirectory)
        {
            if (!System.IO.Path.IsPathRooted(fromAbsolutePath))
                return fromAbsolutePath;  // we can't make a relative if it's not rooted(C:\)  so we'll assume we already have a relative path.

            if (!toDirectory[toDirectory.Length - 1].Equals('\\'))
                toDirectory += "\\";

            Uri from = new Uri(fromAbsolutePath);
            Uri to = new Uri(toDirectory);

            Uri relativeUri = to.MakeRelativeUri(from);
            return relativeUri.ToString();
        }


        public static Property SetValue(this Properties props, string name, object value)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (Property p in props)
            {
                if (p.Name == name)
                {
                    p.Value = value;
                    return p;
                }
            }
            return null;
        }
        public static string AssemblyDirectory()
        {
            string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            return System.IO.Path.GetDirectoryName(path);
        }
        public static string GetDefaultNamespace(this Project project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            return project.Properties.Item("DefaultNamespace").Value.ToString();
        }
        public static string GetProjectDirectory(this Project project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            return System.IO.Path.GetDirectoryName(project.FullName);
        }
        public static Project GetSelectedProject(this EnvDTE80.DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Array projects = dte.ActiveSolutionProjects as Array;
            if (projects?.Length > 0)
            {
                return projects.GetValue(0) as Project;
            }
            return null;
        }
        public static Project GetSelectedProject(this DTE dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Array projects = dte.ActiveSolutionProjects as Array;
            if (projects?.Length > 0)
            {
                return projects.GetValue(0) as Project;
            }
            return null;
        }
        
        public static System.Windows.Window GetMainWindow(this EnvDTE80.DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (dte == null)
            {
                throw new ArgumentNullException("dte");
            }

            var hwndMainWindow = (IntPtr)dte.MainWindow.HWnd;
            if (hwndMainWindow == IntPtr.Zero)
            {
                throw new NullReferenceException("DTE.MainWindow.HWnd is null.");
            }

            var hwndSource = System.Windows.Interop.HwndSource.FromHwnd(hwndMainWindow);
            if (hwndSource == null)
            {
                throw new NullReferenceException("HwndSource for DTE.MainWindow is null.");
            }

            return (System.Windows.Window)hwndSource.RootVisual;
        }
        public static System.Windows.Forms.Screen CurrentScreen(this Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            return System.Windows.Forms.Screen.FromPoint(new System.Drawing.Point(window.Left, window.Top));
        }
    }
}
