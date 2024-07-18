using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestLeft
    {
        [TestMethod]
        public void TestLeftSuccess()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.Left(txt, 15);
            Assert.AreEqual("This is a beaut", v);
        }

        [TestMethod]
        public void TestLeftNULL()
        {
            string txt = "day!";
            Action act = () => StringFunctions.Left(txt, null);
            Assert.ThrowsException<ArgumentException>(act);
        }

        [TestMethod]
        public void TestLeftZeroLen()
        {
            string txt = "day!";
            object v = StringFunctions.Left(txt, 0);
            Assert.AreEqual("", v);
        }
        [TestMethod]
        public void TestLeftGreaterLength()
        {
            string txt = "day!";
            object v = StringFunctions.Left(txt, 1247);
            Assert.AreEqual("day!", v);
        }
        [TestMethod]
        public void TestLeftInt()
        {
            object txt = 1245425635;
            object v = StringFunctions.Left(txt, 2);
            Assert.AreEqual("12", v);
        }
    }
}
