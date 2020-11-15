using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Collections.ObjectModel;

// ReSharper disable InconsistentNaming
namespace CrmCodeGenerator.VSPackage.Model
{
    public class Settings : INotifyPropertyChanged
    {
        public Settings()
        {
            EntityList = new ObservableCollection<string>();
            EntitiesSelected = new ObservableCollection<string>();

            Dirty = false;
        }

        #region boiler-plate INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) 
                return false;

            field = value;
            Dirty = true;
            OnPropertyChanged(propertyName);
            return true;
        }
        #endregion

        private bool _UseSSL;
        private bool _UseIFD;
        private bool _UseOnline;
        private bool _UseOffice365;
        private string _OutputPath;
        private string _Namespace;
        private string _EntitiesToIncludeString;
        private string _CrmOrg;
        private string _Password;
        private string _Username;
        private string _Domain;
        private string _Template;
        private string _T4Path;
        private bool _IncludeNonStandard;
        private bool _IncludeUnpublish;

        private string _ProjectName;
        public string ProjectName
        {
            get => _ProjectName;
            set => SetField(ref _ProjectName, value);
        }
        public string T4Path
        {
            get => _T4Path;
            set => SetField(ref _T4Path, value);
        }
        public string Template
        {
            get => _Template;
            set
            {
                SetField(ref _Template, value);
                NewTemplate = !System.IO.File.Exists(System.IO.Path.Combine(_Folder, _Template));
            }
        }
        private string _Folder = "";
        public string Folder
        {
            get => _Folder;
            set => SetField(ref _Folder, value);
        }

        private bool _NewTemplate;
        public bool NewTemplate
        {
            get => _NewTemplate;
            set => SetField(ref _NewTemplate, value);
        }
        public string OutputPath
        {
            get => _OutputPath;
            set => SetField(ref _OutputPath, value);
        }

        public string Domain
        {
            get => _Domain;
            set => SetField(ref _Domain, value);
        }
        public string Username
        {
            get => _Username;
            set => SetField(ref _Username, value);
        }
        public string Password
        {
            get => _Password;
            set => SetField(ref _Password, value);
        }
        public string CrmOrg
        {
            get => _CrmOrg;
            set => SetField(ref _CrmOrg, value);
        }


        private ObservableCollection<string> _OnLineServers = new ObservableCollection<string>();
        public ObservableCollection<string> OnLineServers
        {
            get => _OnLineServers;
            set => SetField(ref _OnLineServers, value);
        }
        //private string _OnlineServer;
        //public string OnlineServer
        //{
        //    get
        //    {
        //        return _OnlineServer;
        //    }
        //    set
        //    {
        //        SetField(ref _OnlineServer, value);
        //    }
        //}
        private string _ServerName = "";
        public string ServerName
        {
            get => _ServerName;
            set => SetField(ref _ServerName, value);
        }
        private string _ServerPort = "";
        public string ServerPort
        {
            get => UseOnline || UseOffice365 ? "" : _ServerPort;
            set => SetField(ref _ServerPort, value);
        }
        private string _HomeRealm = "";
        public string HomeRealm
        {
            get => _HomeRealm;
            set => SetField(ref _HomeRealm, value);
        }


        private ObservableCollection<string> _OrgList = new ObservableCollection<string>();
        public ObservableCollection<string> OrgList
        {
            get => _OrgList;
            set => SetField(ref _OrgList, value);
        }


        private ObservableCollection<string> _TemplateList = new ObservableCollection<string>();
        public ObservableCollection<string> TemplateList
        {
            get => _TemplateList;
            set => SetField(ref _TemplateList, value);
        }

        public string EntitiesToIncludeString
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                foreach (var value in _EntitiesSelected)
                {
                    if (sb.Length != 0)
                        sb.Append(',');
                    sb.Append(value);
                }
                return sb.ToString();
            }
            set
            {
                var newList = new ObservableCollection<string>();
                var split = value.Split(',').Select(p => p.Trim()).ToList();
                foreach (var s in split)
                {
                    newList.Add(s);
                    if (!_EntityList.Contains(s))
                        _EntityList.Add(s);
                }
                EntitiesSelected = newList;
                SetField(ref _EntitiesToIncludeString, value);
                OnPropertyChanged("EnableExclude");
            }
        }

        private ObservableCollection<string> _EntityList;
        public ObservableCollection<string> EntityList
        {
            get => _EntityList;
            set => SetField(ref _EntityList, value);
        }

        private ObservableCollection<string> _EntitiesSelected;
        public ObservableCollection<string> EntitiesSelected
        {
            get => _EntitiesSelected;
            set => SetField(ref _EntitiesSelected, value);
        }
        public bool IsReadOnly =>_MappingSettings != null;

             private MappingSettings _MappingSettings;
        public MappingSettings MappingSettings
        {
            get => _MappingSettings;
            set => SetField(ref _MappingSettings, value);
        }


        public string Namespace
        {
            get => _Namespace;
            set => SetField(ref _Namespace, value);
        }

        public bool Dirty { get; set; }
        
        public bool IncludeNonStandard
        {
            get => _IncludeNonStandard;
            set => SetField(ref _IncludeNonStandard, value);
        }
        public bool IncludeUnpublish
        {
            get => _IncludeUnpublish;
            set => SetField(ref _IncludeUnpublish, value);
        }
        public bool UseSSL
        {
            get => _UseSSL;
            set
            {
                if (SetField(ref _UseSSL, value))
                {
                    ReEvalReadOnly();
                }
            }
        }
        public bool UseIFD
        {
            get => _UseIFD;
            set
            {
                if (SetField(ref _UseIFD, value) == false)
                    return;

                if (value)
                {
                    UseOnline = false;
                    UseOffice365 = false;
                    UseSSL = true;
                    UseWindowsAuth = false;
                }
                ReEvalReadOnly();
            }
        }
        public bool UseOnline
        {
            get => _UseOnline;
            set
            {
                if (SetField(ref _UseOnline, value) == false)
                    return;

                if (value)
                {
                    UseIFD = false;
                    UseOffice365 = true;
                    UseSSL = true;
                    UseWindowsAuth = false;
                }
                else
                {
                    UseOffice365 = false;
                }
                ReEvalReadOnly();
            }
        }
        public bool UseOffice365
        {
            get => _UseOffice365;
            set
            {
                if (SetField(ref _UseOffice365, value))
                {
                    if (value)
                    {
                        UseIFD = false;
                        UseOnline = true;
                        UseSSL = true;
                        UseWindowsAuth = false;
                    }
                    ReEvalReadOnly();
                }
            }
        }
        private bool _UseWindowsAuth;
        public bool UseWindowsAuth
        {
            get => _UseWindowsAuth;
            set
            {
                SetField(ref _UseWindowsAuth, value);
                ReEvalReadOnly();
            }
        }
        
        #region Read Only Properties
        private void ReEvalReadOnly()
        {
            OnPropertyChanged("NeedServer");
            OnPropertyChanged("NeedOnlineServer");
            OnPropertyChanged("NeedServerPort");
            OnPropertyChanged("NeedHomeRealm");
            OnPropertyChanged("NeedCredentials");
            OnPropertyChanged("CanUseWindowsAuth");
            OnPropertyChanged("CanUseSSL");
        }
        public bool NeedServer => !(UseOnline || UseOffice365);

        public bool NeedOnlineServer => (UseOnline || UseOffice365);

        public bool NeedServerPort => !(UseOffice365 || UseOnline);

        public bool NeedHomeRealm => !(UseIFD || UseOffice365 || UseOnline);

        public bool NeedCredentials => !UseWindowsAuth;

        public bool CanUseWindowsAuth => !(UseIFD || UseOnline || UseOffice365);

        public bool CanUseSSL => !(UseOnline || UseOffice365 || UseIFD);

        #endregion
        
        public string DiscoveryUrl 
            =>  $"{(UseSSL ? "https" : "http")}://{(UseIFD ? ServerName : UseOffice365 ? "disco." + ServerName : UseOnline ? "dev." + ServerName : ServerName)}:{(ServerPort.Length == 0 ? (UseSSL ? 443 : 80) : int.Parse(ServerPort))}/XRMServices/2011/Discovery.svc";
       
        public bool IsActive { get; set; }
    }

    public class MappingSettings
    {
        public Dictionary<string, EntityMappingSetting> Entities;
    }
    public class EntityMappingSetting
    {
        public string CodeName;
        public Dictionary<string, string> Attributes;
    }
}
