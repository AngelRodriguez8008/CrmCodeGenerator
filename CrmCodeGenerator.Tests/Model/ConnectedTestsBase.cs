#region Version
// Last Change: Rodriguez Mustelier Angel (rodang) - 2019-02-13 14:06
// 
#endregion
using System;
using System.Configuration;
using CrmCodeGenerator.VSPackage.Xrm;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace CrmCodeGenerator.Tests
{
    public class ConnectedTestsBase
    {
        // ReSharper disable InconsistentNaming
        protected readonly string Url = ConfigurationManager.AppSettings["CrmUrl"];
        protected readonly string UserName = ConfigurationManager.AppSettings["CrmUser"];
        protected readonly string Password = ConfigurationManager.AppSettings["CrmPassword"];
        protected readonly string Domain = ConfigurationManager.AppSettings["CrmDomain"];
        protected readonly string Org = ConfigurationManager.AppSettings["CrmOrg"];

        private OrganizationServiceProxy _service;
        protected IOrganizationService service => _service ?? CreateService();
        // ReSharper restore InconsistentNaming

        protected OrganizationServiceProxy CreateService() => QuickConnection.Connect(Url, Domain, UserName, Password, Org);

        protected OrganizationServiceProxy CreateService(string otherOrganization) =>  QuickConnection.Connect(Url, Domain, UserName, Password, otherOrganization);
        
        public virtual void TestInit()
        {
            _service = CreateService();
        }

        public virtual void TestFinish()
        {
            _service?.Dispose();
        }

        public Guid? WhoAmI() => service.WhoAmI();
    }
}