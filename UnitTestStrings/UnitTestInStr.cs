using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestInStr
    {
        [TestMethod]
        public void TestFullWithTextComparison()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStr(1, txt, "t", 1);
            Assert.AreEqual(1,v);
        }
        [TestMethod]
        public void TestFullWithBinaryComparison()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStr(1, txt, "t", 0);
            Assert.AreEqual(15,v);
        }
        [TestMethod]
        public void TestFullException()
        {
            
            string txt = "This is a beautiful day!";

            Action act = () => StringFunctions.InStr("xdd", txt, "t", 0);

            Assert.ThrowsException<FormatException>(act);
        }
        [TestMethod]
        public void TestThreeParams()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStr(1, txt, "i");
            Assert.AreEqual(3, v);
        }
        [TestMethod]
        public void TestShortFunc()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStr(txt, "beautiful");
            Assert.AreEqual(11, v);
        }
        [TestMethod]
        public void TestShortFunc1()
        {
            string txt = "bea";
            object v = StringFunctions.InStr(txt, "beautiful");
            Assert.AreEqual(0, v);
        }
        [TestMethod]
        public void TestShortFunc2()
        {
            string txt = "Get a nibble of these soft French rolls, and drink some tea";
            object v = StringFunctions.InStr(txt, "beautiful");
            Assert.AreEqual(0, v);
        }
        [TestMethod]
        public void TestFirstStringEmpty()
        {
            string txt = "";
            object v = StringFunctions.InStr(txt, "beautiful");
            Assert.AreEqual(0, v);
        }
        [TestMethod]
        public void TestSecondStringEmpty()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStr(2,txt, "");
            Assert.AreEqual(2, v);
        }
        [TestMethod]
        public void TestStartGreaterThanString1()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStr(3456789, txt, "beautiful");
            Assert.AreEqual(0, v);
        }
        [TestMethod]
        public void TestFirstStringNUll()
        {
            string txt = null;
            Action act = () => StringFunctions.InStr(txt, "beautiful");
            Assert.ThrowsException<ArgumentException>(act);
        }
    }
}
