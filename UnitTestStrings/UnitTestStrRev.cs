using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestStrRev
    {
        [TestMethod]
        public void TestReverse()
        {
            object rev = StringFunctions.StrReverse("Hello world");
            Assert.AreEqual("dlrow olleH",rev);
        }
        [TestMethod]
        public void TestReverseInt()
        {
            object rev = StringFunctions.StrReverse(123+322);
            Assert.AreEqual("544", rev);
        }
        [TestMethod]
        public void TestReverseBool()
        {
            object rev = StringFunctions.StrReverse(true);
            Assert.AreEqual("eurT", rev);
        }
        [TestMethod]
        public void TestReverseArray()
        {
            char[] charray = "abc".ToCharArray();
            Action act = () => StringFunctions.StrReverse(charray);
            Assert.ThrowsException<ArgumentException>(act);
        }
        [TestMethod]
        public void TestReverseDouble()
        {
            object rev = StringFunctions.StrReverse(0.15);
            Assert.AreEqual("51,0", rev);
        }

    }
}

