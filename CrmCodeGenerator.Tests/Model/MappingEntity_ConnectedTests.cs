using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using CrmCodeGenerator.VSPackage;
using CrmCodeGenerator.VSPackage.Model;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CrmCodeGenerator.VSPackage.Xrm;
using Microsoft.Xrm.Sdk.Metadata;

namespace CrmCodeGenerator.Tests
{

    [TestClass]
    public class MappingEntity_ConnectedTests : ConnectedTestsBase
    {
        [TestInitialize] public override void TestInit() => base.TestInit();
        [TestCleanup] public override void TestFinish() => base.TestFinish();

        [TestMethod]
        public void CheckConnection()
        {
            Assert.IsFalse(WhoAmI().IsNullOrEmpty());
        }

        [TestMethod]
        public void Get_Account_Metadata_StateAttribute()
        {
            var metadata = service.GetEntityMetadata("account");
            Assert.IsNotNull(metadata);

            var state = metadata.Attributes.FirstOrDefault(a => a.LogicalName == "statecode");
            Assert.IsNotNull(state);
        }
        
        
        [TestMethod]
        public void GerOrganizationDetails()
        {
            Console.WriteLine(Url);
            var organizations = QuickConnection.GerOrganizationDetails(Url, Domain, UserName, Password);

            Assert.IsNotNull(organizations);
            Assert.IsTrue(organizations.Length > 0);
            foreach (var organization in organizations)
            {
               Console.WriteLine(organization.UrlName);
            }
        }

        [TestMethod]
        public void GetOrganizationNames()
        {
            Settings settings = new Settings
            {
                UseOnline = true,
                UseSSL = true,
                UseOffice365 = true,
                ServerName = "crm4.dynamics.com",
                Username = "XXXXXXX@XXXXXX.XXXX",
                Password = "XXXXXXX",
                Domain = null
            };
            Console.WriteLine(settings.DiscoveryUrl);
            var organizationNames = QuickConnection.GetOrganizationNames(settings);

            Assert.IsNotNull(organizationNames);
            Assert.IsTrue(organizationNames.Count > 0);
            foreach (var organization in organizationNames)
            {
                Console.WriteLine(organization);
            }
        }

        [TestMethod]
        public void Get_Process_Metadata_StateAttribute()
        {
            var metadata = service.GetEntityMetadata("wag_creditriskprocess");
            Assert.IsNotNull(metadata);

            var state = metadata.Attributes.FirstOrDefault(a => a.LogicalName == "statecode");
            Assert.IsNotNull(state);
            Assert.IsTrue(state is StateAttributeMetadata);
        }

        
        [TestMethod]
        public void Get_Process_Metadata_ActiveStageIdAttribute()
        {
            var metadata = service.GetEntityMetadata("wag_creditriskprocess");
            Assert.IsNotNull(metadata);

            var activeStageId = metadata.Attributes.FirstOrDefault(a => a.LogicalName == "activestageid");
            Assert.IsNotNull(activeStageId);
        }  
        
        [TestMethod]
        public void Get_ProcessInstance_Metadata_StateAttribute()
        {
            var metadata = service.GetEntityMetadata("businessprocessflowinstance");
            Assert.IsNotNull(metadata);

            var state = metadata.Attributes.FirstOrDefault(a => a.LogicalName == "statecode");
            Assert.IsNotNull(state);
        }
        
        [TestMethod]
        public void Get_ProcessInstance_Metadata_ActiveStageIdAttribute()
        {
            var metadata = service.GetEntityMetadata("businessprocessflowinstance");
            Assert.IsNotNull(metadata);

            var activeStageId = metadata.Attributes.FirstOrDefault(a => a.LogicalName == "activestageid");
            Assert.IsNotNull(activeStageId);
        }

        [TestMethod]
        public void GetAllMetadataInclude_ProcessInstance()
        {
            var metadatas = service.GetAllEntitiesMetadata();
            Assert.IsNotNull(metadatas);

            var pi = metadatas.FirstOrDefault(e => e.LogicalName == "businessprocessflowinstance");
            Assert.IsNotNull(pi);
        }
        

        [TestMethod]
        public void GetEntitiesMetadataInclude_ProcessInstance()
        {
            IEnumerable<string> selectedEntities = new []{"account", "businessprocessflowinstance"};
            var metadatas = service.GetEntitiesMetadata(selectedEntities);
            Assert.IsNotNull(metadatas);

            var pi = metadatas.FirstOrDefault(e => e.LogicalName == "businessprocessflowinstance");
            Assert.IsNotNull(pi);
        }

        
        [TestMethod]
        public void GetEntitiesMetadataInclude_GetSelected()
        {
            IEnumerable<string> selectedEntities = new []{"account", "businessprocessflowinstance"};
            var metadatas = service.GetEntitiesMetadata(selectedEntities);
            Assert.IsNotNull(metadatas);

            var mappingSettings = new MappingSettings
            {
                Entities = new Dictionary<string, EntityMappingSetting>{ 
                    {"account", new EntityMappingSetting{
                        CodeName = "account", 
                        Attributes = new Dictionary<string, string>{{"accountid","Id"}}}
                    },
                    {"businessprocessflowinstance", new EntityMappingSetting
                    {
                        CodeName = "ProcessInstance", 
                        Attributes = new Dictionary<string, string> {{"processstageid","ProcessStageId"}}}
                    }
                }
            };

            Settings settings = new Settings
            {
                EntitiesSelected = new ObservableCollection<string>(selectedEntities),
                IncludeNonStandard = false,
                MappingSettings = mappingSettings
            };

            var mapper = new Mapper(settings);
            var selected = mapper.GetSelectedEntities(metadatas);
            var pi = selected.FirstOrDefault(e => e.LogicalName == "businessprocessflowinstance");
            Assert.IsNotNull(pi);
        }
    }
}