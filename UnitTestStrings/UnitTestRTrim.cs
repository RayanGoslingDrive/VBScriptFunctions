using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestRTrim
    {
        [TestMethod]
        public void TestRTrimSuccess()
        {
            object fname = " Jack ";
            object v = "Hello" + StringFunctions.RTrim(fname) + "and welcome.";
            Assert.AreEqual("Hello Jackand welcome.", v);
        }

        [TestMethod]
        public void TestRTrimTwoWords()
        {
            object fname = " Ja ck ";
            object v = "Hello" + StringFunctions.RTrim(fname) + "and welcome.";
            Assert.AreEqual("Hello Ja ckand welcome.", v);
        }

        [TestMethod]
        public void TestRTrimNoSpaces()
        {
            object fname = "Jack";
            object v = "Hello" + StringFunctions.RTrim(fname) + "and welcome.";
            Assert.AreEqual("HelloJackand welcome.", v);
        }

        [TestMethod]
        public void TestRTrimALotOfSpaces()
        {
            object fname = "                               Jack                        ";
            object v = "Hello" + StringFunctions.RTrim(fname) + "and welcome.";
            Assert.AreEqual("Hello                               Jackand welcome.", v);
        }
        [TestMethod]
        public void TestRTrimInt()
        {
            object fname = 1245425635;
            object v = StringFunctions.RTrim(          1245425635             );
            Assert.AreEqual("1245425635", v);
        }
        [TestMethod]
        public void TestRTrimNull()
        {
            object fname = null;
            Action act = () => StringFunctions.RTrim(fname);
            Assert.ThrowsException<ArgumentException>(act);
        }
    }
}
