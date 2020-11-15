using CrmCodeGenerator.VSPackage.Helpers;
using EnvDTE80;

namespace CrmCodeGenerator.VSPackage
{
    public class Configuration
    {

        #region Singleton
        private static Configuration _instance;
        private static readonly object SyncLock = new object();
        public static Configuration Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (SyncLock)
                    {
                        if (_instance == null)
                        {
                            _instance = new Configuration();
                        }
                    }
                }
                return _instance;
            }
        }
        #endregion

        
        public Configuration()
        {
            Settings = new Model.Settings
            {
                ServerName = "crm.dynamics.com",
                ProjectName = "",
                Domain = "",
                T4Path = System.IO.Path.Combine(DteHelper.AssemblyDirectory(), @"Resources\Templates\CrmSvcUtil.tt"),
                Template = "",
                CrmOrg = "DEV-CRM",
                EntitiesToIncludeString = "account,contact,lead,opportunity,systemuser",
                OutputPath = "",
                Username = "@XXXXX.onmicrosoft.com",
                Password = "",
                Namespace = "",
                Dirty = false
            };
        }
        public Model.Settings Settings { get; set; }
        public DTE2 DTE { get; set; }
    }
}
