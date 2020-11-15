using Microsoft.VisualStudio.TestTools.UnitTesting;
using CrmCodeGenerator.VSPackage.Helpers;

namespace CrmCodeGenerator.Tests
{
    [TestClass]
    public class Naming_Tests
    {
        [TestMethod]
        public void GetEnumItemName_AllUpper()
        {
            var result = Naming.GetEnumItemName("ALL");
            Assert.AreEqual("All", result);
        }

        [TestMethod]
        public void GetEnumItemName_AllUpper_MinusSeparator()
        {
            var result = Naming.GetEnumItemName("ALL-UPPER");
            Assert.AreEqual("AllUpper", result);
        }

        [TestMethod]
        public void GetEnumItemName_AllUpper_BlankSeparator()
        {
            var result = Naming.GetEnumItemName("ALL UPPER");
            Assert.AreEqual("AllUpper", result);
        }
        
        [TestMethod]
        public void GetEnumItemName_SentenceCase_BlankSeparator()
        {
            var result = Naming.GetEnumItemName("Sentence case");
            Assert.AreEqual("SentenceCase", result);
        }

        
        [TestMethod]
        public void GetEnumItemName_LowerCase_BlankSeparator()
        {
            var result = Naming.GetEnumItemName("sentence case");
            Assert.AreEqual("SentenceCase", result);
        }
    }
}