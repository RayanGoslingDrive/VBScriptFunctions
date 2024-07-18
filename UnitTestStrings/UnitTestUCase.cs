using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestUCase
    {
        [TestMethod]
        public void TestUCaseSuccess()
        {
            string txt = "qwertyuiop";
            object v = StringFunctions.UCase(txt);
            Assert.AreEqual("QWERTYUIOP", v);
        }
        [TestMethod]
        public void TestUCaseInt()
        {
            object v = StringFunctions.UCase(123456789);
            Assert.AreEqual("123456789", v);
        }
        [TestMethod]
        public void TestUCaseBOOL()
        {
            object v = StringFunctions.UCase(true);
            Assert.AreEqual("TRUE", v);
        }
    }
}
