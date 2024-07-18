using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestStrring
    {
        [TestMethod]
        public void TestMethod1()
        {
            string txt = "qwertyuiop";
            object v = StringFunctions.String(10,txt);
            Assert.AreEqual("qqqqqqqqqq", v);
        }
        [TestMethod]
        public void TestMethod2()
        {
            object v = StringFunctions.String(10, 64);
            Assert.AreEqual("@@@@@@@@@@", v);
        }
    }
}
