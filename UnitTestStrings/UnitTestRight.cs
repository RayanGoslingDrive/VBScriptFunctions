using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestRight
    {
        [TestMethod]
        public void TestRightSuccess()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.Right(txt, 10);
            Assert.AreEqual("tiful day!", v);
        }
        [TestMethod]
        public void TestRightNULLException()
        {
            string txt = "day!";
            Action act = () => StringFunctions.Right(txt, null);
            Assert.ThrowsException<ArgumentException>(act);
        }

        [TestMethod]
        public void TestRightZeroLen()
        {
            string txt = "day!";
            object v = StringFunctions.Right(txt, 0);
            Assert.AreEqual("", v);
        }
        [TestMethod]
        public void TestRightGreaterLength()
        {
            string txt = "day!";
            object v = StringFunctions.Right(txt, 1245);
            Assert.AreEqual("day!", v);
        }
        [TestMethod]
        public void TestRightInt()
        {
            object txt = 1245425635;
            object v = StringFunctions.Right(txt, 2);
            Assert.AreEqual("35", v);
        }
    }
}
