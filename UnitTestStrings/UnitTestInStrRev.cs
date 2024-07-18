using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestInStrRev
    {
        [TestMethod]
        public void TestFullWithTextComparison()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStrRev(txt, "T",-1,1);
            Assert.AreEqual(15,v);
        }
        [TestMethod]
        public void TestFullWithBinaryComparison()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStrRev(txt, "T", -1,0);
            Assert.AreEqual(1,v);
        }
        [TestMethod]
        public void TestThreeParams()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStrRev(txt, "i",-1);
            Assert.AreEqual(16, v);
        }
        [TestMethod]
        public void TestThreeParams2()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStrRev(txt, "i",7);
            Assert.AreEqual(6, v);
        }
        [TestMethod]
        public void TestFullException()
        {
            
            string txt = "This is a beautiful day!";

            Action act = () => StringFunctions.InStrRev( txt, "t","xdd", 0);

            Assert.ThrowsException<FormatException>(act);
        }
        
        [TestMethod]
        public void TestShortFunc()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStrRev(txt, "beautiful");
            Assert.AreEqual(11, v);
        }
        [TestMethod]
        public void TestShortFunc1()
        {
            string txt = "bea";
            object v = StringFunctions.InStrRev(txt, "beautiful");
            Assert.AreEqual(0, v);
        }
        [TestMethod]
        public void TestShortFunc2()
        {
            string txt = "Get a nibble of these soft French rolls, and drink some tea";
            object v = StringFunctions.InStrRev(txt, "beautiful");
            Assert.AreEqual(0, v);
        }
        [TestMethod]
        public void TestFirstStringEmpty()
        {
            string txt = "";
            object v = StringFunctions.InStrRev(txt, "beautiful");
            Assert.AreEqual(0, v);
        }
        [TestMethod]
        public void TestSecondStringEmpty()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStrRev(txt, "",2);
            Assert.AreEqual(2, v);
        }
        [TestMethod]
        public void TestStartGreaterThanString1()
        {
            string txt = "This is a beautiful day!";
            object v = StringFunctions.InStrRev( txt, "beautiful",3456789);
            Assert.AreEqual(0, v);
        }
        [TestMethod]
        public void TestFirstStringNUll()
        {
            string txt = null;
            Action act = () => StringFunctions.InStrRev(txt, "beautiful");
            Assert.ThrowsException<ArgumentException>(act);
        }
    }
}
