using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using CrmCodeGenerator.VSPackage.Helpers;
using CrmCodeGenerator.VSPackage.Model;
using Microsoft.VisualStudio.TextTemplating.VSHost;
using Task = System.Threading.Tasks.Task;

namespace CrmCodeGenerator.VSPackage
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(AllowsBackgroundLoading = true, UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    //this causes the class to load when VS starts [ProvideAutoLoad("ADFC4E64-0397-11D1-9F4E-00A0C911004F")]
    //[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string)]
    //[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidCrmCodeGenerator_VSPackagePkgString)]
    [ProvideSolutionProps(_strSolutionPersistanceKey)]
    [ProvideCodeGenerator(typeof(XrmCodeGenerator), XrmCodeGenerator.Name, XrmCodeGenerator.Description, true)]
    [ProvideUIContextRule("69760bd3-80f0-4901-818d-c4656aaa08e9", // Must match the GUID in the .vsct file
        name: "UI Context",
        expression: "tt", // This will make the button only show on .tt files (Ex: js | css | html)
        termNames: new[] { "tt" },
        termValues: new[] { "HierSingleSelectionName:.tt$" })]
    public sealed class CrmCodeGenerator_VsPackagePackage : AsyncPackage, IVsPersistSolutionProps, IVsSolutionEvents3
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public CrmCodeGenerator_VsPackagePackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", ToString()));
        }
        
        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members
            
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
         //   Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering InitializeAsync() of: {0}", this.ToString()));

            // Request any services while on the background thread
            var commandService = await GetServiceAsync((typeof(IMenuCommandService))) as IMenuCommandService;

            // Switch to the main thread before initializing the AddTemplateButton command
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // Initialize the AddTemplateButton command and pass it the commandService
            AddTemplateButton.Initialize(this, commandService);

            // Initialize the ApplyCustomTool Command
            await ApplyCustomTool.InitializeAsync(this);

            Configuration.Instance.DTE = await GetServiceAsync(typeof(SDTE)) as EnvDTE80.DTE2;
            AdviseSolutionEvents();
        }

        protected override void Dispose(bool disposing)
        {
            UnadviseSolutionEvents();

            base.Dispose(disposing);
        }

        private IVsSolution solution;
        private uint handleCookie;
        private void AdviseSolutionEvents()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            UnadviseSolutionEvents();

            solution = GetService(typeof(SVsSolution)) as IVsSolution;

            solution?.AdviseSolutionEvents(this, out handleCookie);
        }

        private void UnadviseSolutionEvents()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (solution == null)
                return;
            
            if (handleCookie != uint.MaxValue)
            {
                solution.UnadviseSolutionEvents(handleCookie);
                handleCookie = uint.MaxValue;
            }

            solution = null;
        }

        #endregion

        #region IVsPersistSolutionProps Implementation Code

        private readonly Settings settings = Configuration.Instance.Settings;

        private const string _strSolutionPersistanceKey = "CrmCodeGeneration";
        private const string _strCrmUrl = "CrmUrl";
        private const string _strUseSSL = "UseSSL";
        private const string _strUseIFD = "UseIFD";
        private const string _strUseOnline = "UseOnline";
        private const string _strUseOffice365 = "UseOffice365";
        private const string _strServerPort = "ServerPort";
        private const string _strServerName = "ServerName";
        private const string _strHomeRealm = "HomeRealm";
        private const string _strUsername = "Username";
        private const string _strPassword = "Password";
        private const string _strDomain = "Domain";
        private const string _strUseWindowsAuth = "WindowsAuthorization";
        private const string _strOrganization = "Organization";
        private const string _strIncludeEntities = "IncludeEntities";
        private const string _strIncludeNonStandard = "IncludeNonStandard";

        #region Solution Properties
        public int QuerySaveSolutionProps(IVsHierarchy pHierarchy, VSQUERYSAVESLNPROPS[] pqsspSave)
        {
            if (pHierarchy != null) // if this contains something, then VS is asking for Solution Properties of a PROJECT,  
                pqsspSave[0] = VSQUERYSAVESLNPROPS.QSP_HasNoProps;
            else
            {
                if (settings.Dirty)
                    pqsspSave[0] = VSQUERYSAVESLNPROPS.QSP_HasDirtyProps;
                else
                    pqsspSave[0] = VSQUERYSAVESLNPROPS.QSP_HasNoDirtyProps;
            }
            return VSConstants.S_OK;
        }
        public int SaveSolutionProps([In] IVsHierarchy pHierarchy, [In] IVsSolutionPersistence pPersistence)
        {
            // This function gets called by the shell after determining the package has dirty props.
            // The package will pass in the key under which it wants to save its properties, 
            // and the IDE will call back on WriteSolutionProps

            // The properties will be saved in the Pre-Load section
            // When the solution will be reopened, the IDE will call our package to load them back before the projects in the solution are actually open
            // This could help if the source control package needs to persist information like projects translation tables, that should be read from the suo file
            // and should be available by the time projects are opened and the shell start calling IVsSccEnlistmentPathTranslation functions.
            ThreadHelper.ThrowIfNotOnUIThread();
            
            if (settings.IsActive && settings.Dirty)
                pPersistence.SavePackageSolutionProps(1, null, this, _strSolutionPersistanceKey);

            settings.Dirty = false;

            return VSConstants.S_OK;
        }
        public int WriteSolutionProps([In] IVsHierarchy pHierarchy, [In] string pszKey, [In] IPropertyBag pPropBag)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            pPropBag.WriteBool(_strUseSSL, settings.UseSSL);
            pPropBag.WriteBool(_strUseIFD, settings.UseIFD);
            pPropBag.WriteBool(_strUseOnline, settings.UseOnline);
            pPropBag.WriteBool(_strUseOffice365, settings.UseOffice365);
            pPropBag.WriteBool(_strUseOffice365, settings.UseOffice365);

            pPropBag.Write(_strServerName, settings.ServerName);
            pPropBag.Write(_strServerPort, settings.ServerPort);
            pPropBag.Write(_strHomeRealm, settings.HomeRealm);

            //pPropBag.Write(_strCrmUrl, settings.CrmSdkUrl);
            pPropBag.Write(_strDomain, settings.Domain);
            pPropBag.Write(_strUseWindowsAuth, settings.UseWindowsAuth.ToString());
            //pPropBag.Write(_strUsername, settings.Username);
            //pPropBag.Write(_strPassword, settings.Password);

            pPropBag.Write(_strOrganization, settings.CrmOrg);
            pPropBag.Write(_strIncludeEntities, settings.EntitiesToIncludeString);
            pPropBag.WriteBool(_strIncludeNonStandard, settings.IncludeNonStandard);
            settings.Dirty = false;

            return VSConstants.S_OK;
        }
        public int ReadSolutionProps(IVsHierarchy pHierarchy, string pszProjectName, string pszProjectMk, string pszKey, int fPreLoad, IPropertyBag pPropBag)
        {
            if (string.Compare(_strSolutionPersistanceKey, pszKey, StringComparison.Ordinal) != 0)
                return VSConstants.S_OK;

            var defaultServer = "crm.dynamics.com";
            var defaultSSL = false;
            var defaultPort = "";
            var defaultOnline = true;


            // This is to convert the earlier settings
            var oldServerSetting = pPropBag.Read(_strCrmUrl, "");
            if (!string.IsNullOrWhiteSpace(oldServerSetting))
            {
                Uri oldServer = new Uri(oldServerSetting);
                defaultServer = oldServer.Host;
                defaultSSL = (oldServer.Scheme == "https");
                defaultPort = oldServer.Port.ToString();
                defaultOnline = (oldServer.Host.ToLower() == "crm.dynamics.com");
            }


            settings.ServerName = pPropBag.Read(_strServerName, defaultServer);
            settings.UseSSL = pPropBag.Read(_strUseSSL, defaultSSL);
            settings.UseIFD = pPropBag.Read(_strUseIFD, false);
            settings.UseOnline = pPropBag.Read(_strUseOnline, defaultOnline);
            settings.UseOffice365 = pPropBag.Read(_strUseOffice365, defaultOnline);
            settings.ServerPort = pPropBag.Read(_strServerPort, defaultPort);
            settings.HomeRealm = pPropBag.Read(_strHomeRealm, "");

            //settings.Username = pPropBag.Read(_strUsername, "");
            //settings.Password = pPropBag.Read(_strPassword, "");
            settings.Domain = pPropBag.Read(_strDomain, "");
            settings.UseWindowsAuth = pPropBag.Read(_strUseWindowsAuth, false);
            settings.IsActive = pPropBag.HasSetting(_strOrganization);
            settings.CrmOrg = pPropBag.Read(_strOrganization, "DEV-CRM");
            settings.EntitiesToIncludeString = pPropBag.Read(_strIncludeEntities, "account, contact, systemuser");

            settings.IncludeNonStandard = pPropBag.Read(_strIncludeNonStandard, false);

            settings.Dirty = false;
            return VSConstants.S_OK;
        }
        #endregion
        #region User Options
        public int LoadUserOptions(IVsSolutionPersistence pPersistence, uint grfLoadOpts)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            pPersistence.LoadPackageUserOpts(this, _strSolutionPersistanceKey + _strUsername);
            pPersistence.LoadPackageUserOpts(this, _strSolutionPersistanceKey + _strPassword);
            return VSConstants.S_OK;
        }
        public int ReadUserOptions(IStream pOptionsStream, string pszKey)
        {
            try
            {
                using (StreamEater wrapper = new StreamEater(pOptionsStream))
                {
                    string value;
                    using (var bReader = new System.IO.BinaryReader(wrapper))
                    {
                        value = bReader.ReadString();
                        using (var aes = new SimpleAES())
                        {
                            value = aes.Decrypt(value);
                        }
                    }

                    switch (pszKey)
                    {
                        case _strSolutionPersistanceKey + _strUsername:
                            settings.Username = value;
                            break;
                        case _strSolutionPersistanceKey + _strPassword:
                            settings.Password = value;
                            break;
                    }
                }
                return VSConstants.S_OK;
            }
            finally
            {
                Marshal.ReleaseComObject(pOptionsStream);
            }
        }

        public int SaveUserOptions(IVsSolutionPersistence pPersistence)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            pPersistence.SavePackageUserOpts(this, _strSolutionPersistanceKey + _strUsername);
            pPersistence.SavePackageUserOpts(this, _strSolutionPersistanceKey + _strPassword);
            return VSConstants.S_OK;
        }

        public int WriteUserOptions(IStream pOptionsStream, string pszKey)
        {
            try
            {
                string value;
                switch (pszKey)
                {
                    case _strSolutionPersistanceKey + _strUsername:
                        value = settings.Username;
                        break;
                    case _strSolutionPersistanceKey + _strPassword:
                        value = settings.Password;
                        break;
                    default:
                        return VSConstants.S_OK;
                }

                using (var aes = new SimpleAES())
                {
                    value = aes.Encrypt(value);
                    using (StreamEater wrapper = new StreamEater(pOptionsStream))
                    {
                        using (var bw = new System.IO.BinaryWriter(wrapper))
                        {
                            bw.Write(value);
                        }
                    }
                }
                return VSConstants.S_OK;
            }
            finally
            {
                Marshal.ReleaseComObject(pOptionsStream);
            }
        }
        #endregion
        public int OnProjectLoadFailure(IVsHierarchy pStubHierarchy, string pszProjectName, string pszProjectMk, string pszKey)
        {
            return VSConstants.S_OK;
        }
        #endregion


        #region SolutionEvents
        public int OnAfterCloseSolution(object pUnkReserved) { settings.IsActive = false; return VSConstants.S_OK; }
        public int OnAfterClosingChildren(IVsHierarchy pHierarchy) { return VSConstants.S_OK; }
        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy) { return VSConstants.S_OK; }
        public int OnAfterMergeSolution(object pUnkReserved) { return VSConstants.S_OK; }
        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded) { return VSConstants.S_OK; }
        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution) { return VSConstants.S_OK; }
        public int OnAfterOpeningChildren(IVsHierarchy pHierarchy) { return VSConstants.S_OK; }
        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved) { return VSConstants.S_OK; }
        public int OnBeforeCloseSolution(object pUnkReserved) { return VSConstants.S_OK; }
        public int OnBeforeClosingChildren(IVsHierarchy pHierarchy) { return VSConstants.S_OK; }
        public int OnBeforeOpeningChildren(IVsHierarchy pHierarchy) { return VSConstants.S_OK; }
        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy) { return VSConstants.S_OK; }
        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel) { return VSConstants.S_OK; }
        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel) { return VSConstants.S_OK; }
        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel) { return VSConstants.S_OK; }
        #endregion
    }
}
