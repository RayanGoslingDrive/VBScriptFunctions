using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestLCase
    {
        [TestMethod]
        public void TestLCaseSuccess()
        {
            string txt = "THIS IS A BEAUTIFUL DAY!";
            object v = StringFunctions.LCase(txt);
            Assert.AreEqual("this is a beautiful day!", v);
        }
        [TestMethod]
        public void TestLCaseInt()
        {
            object txt = 1111113111;
            object v = StringFunctions.LCase(txt);
            Assert.AreEqual("1111113111", v);
        }
    }
}
