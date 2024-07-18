using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestLTrim
    {
        [TestMethod]
        public void TestLTrimSuccess()
        {
            object fname = " Jack ";
            object v = "Hello" + StringFunctions.LTrim(fname) + "and welcome.";
            Assert.AreEqual("HelloJack and welcome.", v);
        }
        [TestMethod]
        public void TestLTrimTwoWords()
        {
            object fname = " Ja ck ";
            object v = "Hello" + StringFunctions.LTrim(fname) + "and welcome.";
            Assert.AreEqual("HelloJa ck and welcome.", v);
        }

        [TestMethod]
        public void TestLTrimNoSpaces()
        {
            object fname = "Jack ";
            object v = "Hello" + StringFunctions.LTrim(fname) + "and welcome.";
            Assert.AreEqual("HelloJack and welcome.", v);
        }

        [TestMethod]
        public void TestLTrimALotOfSpaces()
        {
            object fname = "                               Jack ";
            object v = "Hello" + StringFunctions.LTrim(fname) + "and welcome.";
            Assert.AreEqual("HelloJack and welcome.", v);
        }
        [TestMethod]
        public void TestLTrimArgumentException()
        {
            object fname = 1245425635;
            object v = StringFunctions.LTrim(       1245425635          );
            Assert.AreEqual("1245425635", v);
        }
        [TestMethod]
        public void TestLTrimNullRefException()
        {
            object fname = null;
            Action act = () => StringFunctions.LTrim(fname);
            Assert.ThrowsException<ArgumentException>(act);
        }
    }
}
