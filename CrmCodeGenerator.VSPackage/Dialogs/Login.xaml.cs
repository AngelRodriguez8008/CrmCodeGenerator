using CrmCodeGenerator.VSPackage.Helpers;
using CrmCodeGenerator.VSPackage.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CrmCodeGenerator.VSPackage.Xrm;

namespace CrmCodeGenerator.VSPackage.Dialogs
{
    /// <summary>
    /// Interaction logic for Login.xaml
    /// </summary>
    public partial class Login
    {
        public Mapper Mapper;
        private readonly Settings settings;
        private bool entitiesLoaded = false;

        public bool StillOpen { get; private set; } = true;

        public Login(EnvDTE80.DTE2 dte, Settings settings)
        {
            WifDetector.CheckForWifInstall();
            InitializeComponent();

            var main = dte.GetMainWindow();
            Owner = main;
            //Loaded += delegate  { this.CenterWindow(main); };

            this.settings = settings;
            txtPassword.Password = settings.Password;  // PasswordBox doesn't allow 2 way binding
            DataContext = settings;
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            this.HideMinimizeAndMaximizeButtons();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            settings.IsActive = true;
            UpdateStatus("In order to generate code from this template, you need to provide login credentials for your CRM system", null);
            UpdateStatus("The Discovery URL is the URL to your Discovery Service, you can find this URL in CRM -> Settings -> Customizations -> Developer Resources.  \n    eg " + @"https://dsc.yourdomain.com/XRMServices/2011/Discovery.svc", null);
            if (settings.OrgList.Contains(settings.CrmOrg) == false)
            {
                settings.OrgList.Add(settings.CrmOrg);
            }
            Organization.SelectedItem = settings.CrmOrg;
        }
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            StillOpen = false;
            Close();
        }

        private void RefreshOrgs(object sender, RoutedEventArgs e)
        {
            settings.Password = ((PasswordBox)((Button)sender).CommandParameter).Password;  // PasswordBox doesn't allow 2 way binding, so we have to manually read it
            try
            {
                UpdateStatus("Refreshing Organizations", true);
                List<string> organizationNames = QuickConnection.GetOrganizationNames(settings);
                settings.OrgList = new ObservableCollection<string>(organizationNames);
                UpdateStatus("Organizations Loaded. Please pick one", false);
            }
            catch (Exception ex)
            {
                var error = "[ERROR] " + ex.Message + (ex.InnerException != null ? "\n" + "[ERROR] " + ex.InnerException.Message : "");
                UpdateStatus(error, false);
                UpdateStatus("Unable to refresh organizations, check connection information", false);
            }
        }

        private void EntitiesRefresh_Click(object sender, RoutedEventArgs events)
        {
            settings.Password = ((PasswordBox)((Button)sender).CommandParameter).Password;  // PasswordBox doesn't allow 2 way binding, so we have to manually read it

            UpdateStatus("Refreshing Entities...", true);

            RefreshEntityList();

            UpdateStatus("", false);
        }

        private void RefreshEntityList()
        {
            List<string> allEntities = GetAllEntityNames();
            if (allEntities == null)
                return;

            List<string> entities = settings.IncludeNonStandard
                            ? allEntities
                            : allEntities.Except(EntityHelper.NonStandard).ToList();
      
            entities.Sort();
            
            string origSelection = settings.EntitiesToIncludeString;
            settings.EntityList = new ObservableCollection<string>(entities);
            settings.EntitiesToIncludeString = origSelection;
            entitiesLoaded = true;
        }

        private List<string> GetAllEntityNames()
        {
            try
            {
                using (var service = QuickConnection.Connect(settings))
                {
                    List<string> result = service.GetAllEntityNames(settings.IncludeUnpublish);
                    return result;
                }
            }
            catch (Exception ex)
            {
                var error = "[ERROR] " + ex.Message + (ex.InnerException != null ? "\n" + "[ERROR] " + ex.InnerException.Message : "");
                UpdateStatus(error, false);
                UpdateStatus("Unable to refresh entities, check connection information", false);
                return null;
            }
        }

        private void IncludeNonStandardEntities_Click(object sender, RoutedEventArgs e)
        {
            if (entitiesLoaded)
                RefreshEntityList(); // if we don't have the entire list of entities don't do anything (eg if they haven't entered a username & password)
        }

        private void Logon_Click(object sender, RoutedEventArgs e)
        {
            settings.Password = ((PasswordBox)((Button)sender).CommandParameter).Password;   // PasswordBox doesn't allow 2 way binding, so we have to manually read it
            UpdateStatus("Logging in to CRM...", true);
            try
            {
                using (var service = QuickConnection.Connect(settings))
                {
                    bool success = service.CheckConnection();
                    if (!success)
                    {
                        UpdateStatus("Unable to login to CRM, check to ensure you have the right organization", false);
                        return;
                    }
                }
                UpdateStatus("Mapping entities, this might take a while depending on CRM server/connection speed... ", true);
                Mapper = new Mapper(settings);

                settings.Dirty = true;  //  TODO Because the EntitiesSelected is a collection, the Settings class can't see when an item is added or removed.  when I have more time need to get the observable to bubble up.
                StillOpen = false;
                Close();
            }
            catch (Exception ex)
            {
                var error = "[ERROR] " + ex.Message + (ex.InnerException != null ? "\n" + "[ERROR] " + ex.InnerException.Message : "");
                UpdateStatus(error, false);
                UpdateStatus(ex.StackTrace, false);
                UpdateStatus("Unable to map entities, see error above.", false);
            }
            UpdateStatus("", false);
        }
        private void UpdateStatus(string message, bool? working)
        {
            if (working == true)
            {
                Dispatcher?.BeginInvoke(new Action(() =>
                {
                    Cursor = Cursors.Wait;
                    Inputs.IsEnabled = false;
                }));
            }
            if (working == false)
            {
                Dispatcher?.BeginInvoke(new Action(() =>
                {
                    Cursor = null;
                    Inputs.IsEnabled = true;
                }));
            }

            if (!string.IsNullOrWhiteSpace(message))
            {
                Dispatcher?.BeginInvoke(new Action(() => { Status.Update(message); }));
            }

            System.Windows.Forms.Application.DoEvents();  // Needed to allow the output window to update (also allows the cursor wait and form disable to show up)
        }
    }
}
