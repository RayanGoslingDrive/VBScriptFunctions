using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestMid
    {
        [TestMethod]
        public void TestMidSuccess()
        {
            object txt = "This is a sandwich!";
            object v = StringFunctions.Mid(txt, 1, 1);
            Assert.AreEqual("T", v);
        }
        [TestMethod]
        public void TestMidSuccess2()
        {
            object txt = "This is a sandwich!";
            object v = StringFunctions.Mid(txt, 1, 9);
            Assert.AreEqual("This is a", v);
        }
        [TestMethod]
        public void TestMidShort()
        {
            object txt = "This is a sandwich!";
            object v = StringFunctions.Mid(txt, 4);
            Assert.AreEqual("s is a sandwich!", v);
        }
        [TestMethod]
        public void TestMidGreaterStart()
        {
            object txt = "This is a sandwich!";
            object v = StringFunctions.Mid(txt, 15678, 1);
            Assert.AreEqual("", v);
        }
        [TestMethod]
        public void TestMidZeroStart()
        {
            object txt = "This is a sandwich!";
            Action act = () => StringFunctions.Mid(txt, 0, 1);
            Assert.ThrowsException<ArgumentOutOfRangeException>(act);
        }
        [TestMethod]
        public void TestMidnullStart()
        {
            object txt = "This is a sandwich!";
            Action act = () => StringFunctions.Mid(txt, null);
            Assert.ThrowsException<ArgumentException>(act);
        }
        [TestMethod]
        public void TestMidGreaterLength()
        {
            object txt = "This is a sandwich!";
            object v = StringFunctions.Mid(txt, 1,56789);
            Assert.AreEqual(txt, v);
        }
        [TestMethod]
        public void TestMidZeroLength()
        {
            object txt = "This is a sandwich!";
            object v = StringFunctions.Mid(txt, 1, 0);
            Assert.AreEqual("", v);
        }

    }
}
