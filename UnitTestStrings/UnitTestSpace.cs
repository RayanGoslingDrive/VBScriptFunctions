using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestSpace
    {
        [TestMethod]
        public void TestSpaceTen()
        {
            object txt = StringFunctions.Space(10);
            Assert.AreEqual("          ", txt);
        }
        [TestMethod]
        public void TestSpaceTenString()
        {
            object txt = StringFunctions.Space("10");
            Assert.AreEqual("          ", txt);

        }
        [TestMethod]
        public void TestSpaceNull()
        {
            Action acct = () => StringFunctions.Space(null);
            Assert.ThrowsException<ArgumentException>(() => acct());

        }
    }
}
