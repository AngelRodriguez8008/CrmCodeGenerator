using Microsoft.VisualStudio.TestTools.UnitTesting;
using CrmCodeGenerator.VSPackage.Model;
using CrmCodeGenerator.VSPackage;

namespace CrmCodeGenerator.Tests
{
    [TestClass]
    public class Settings_Tests
    {
        [TestMethod]
        public void Deserialize_MappingSettings()
        {
            string json = @"{
""Entities"": {
    ""account"": {
      ""CodeName"": ""Account"",
      ""Attributes"": {
        ""name"": ""Name""
      }
    },
    ""activityparty"" : {
			""CodeName"": ""ActivityParty"",
			""Attributes"": {    
				""activityid"" : ""ActivityId"",
				""activityidName"" : ""ActivityIdName"",
				""activitypartyid"" : ""ActivityPartyId"",
				""addressused"" : ""AddressUsed"",
				""addressusedemailcolumnnumber"" : ""AddressUsedEmailColumnNumber"",
				""donotemail"" : ""DoNotEmail"",
				""donotfax"" : ""DoNotFax"",
				""donotphone"" : ""DoNotPhone"",
				""donotpostalmail"" : ""DoNotPostalMail"",
				""effort"" : ""Effort"",
				""exchangeentryid"" : ""ExchangeEntryId"",			
				""versionnumber"" : ""VersionNumber""
          }
	    },
    }
}";
            var mapping = VSPackage.XrmCodeGenerator.Deserialize<MappingSettings>(json);
            Assert.IsNotNull(mapping);
            Assert.IsNotNull(mapping.Entities);
            Assert.AreEqual(2, mapping.Entities.Count);
            Assert.IsNotNull(mapping.Entities["account"]);
            Assert.IsNotNull(mapping.Entities["account"].Attributes["name"]);
        }
    }
}