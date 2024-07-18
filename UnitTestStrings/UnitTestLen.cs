using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestLen
    {
        [TestMethod]
        public void TestLenSuccess()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.Len(txt);
            Assert.AreEqual(24,v);
        }

        [TestMethod]
        public void TestLenNULLException()
        {
            string txt = null;
            Action act = () => StringFunctions.Len(txt);
            Assert.ThrowsException<ArgumentException>(act);
        }

        [TestMethod]
        public void TestLenZeroLen()
        {
            string txt = "";
            object v = StringFunctions.Len(txt);
            Assert.AreEqual(0, v);
        }
        [TestMethod]
        public void TestLenArgumentException()
        {
            object txt = 1245425635;
            object v = StringFunctions.Len(txt);
            Assert.AreEqual(10, v);
        }
    }
}
