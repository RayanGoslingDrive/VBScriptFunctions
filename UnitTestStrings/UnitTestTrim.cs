using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestTrim
    {
        [TestMethod]
        public void TestTrimSuccess()
        {
            object fname = " Jack ";
            object v = "Hello" + StringFunctions.Trim(fname) + "and welcome.";
            Assert.AreEqual("HelloJackand welcome.", v);
        }

        [TestMethod]
        public void TestTrimTwoWords()
        {
            object fname = " Ja ck ";
            object v = "Hello" + StringFunctions.Trim(fname) + "and welcome.";
            Assert.AreEqual("HelloJa ckand welcome.", v);
        }

        [TestMethod]
        public void TestTrimNoSpaces()
        {
            object fname = "Jack";
            object v = "Hello" + StringFunctions.Trim(fname) + "and welcome.";
            Assert.AreEqual("HelloJackand welcome.", v);
        }

        [TestMethod]
        public void TestTrimALotOfSpaces()
        {
            object fname = "                               Jack                        ";
            object v = "Hello" + StringFunctions.Trim(fname) + "and welcome.";
            Assert.AreEqual("HelloJackand welcome.", v);
        }
        [TestMethod]
        public void TestTrimArgumentException()
        {
            object fname = 1245425635;
            object v = StringFunctions.Trim(        1245425635          );
            Assert.AreEqual("1245425635", v);
        }
        [TestMethod]
        public void TestTrimNullRefException()
        {
            object fname = null;
            Action act = () => StringFunctions.Trim(fname);
            Assert.ThrowsException<ArgumentException>(act);
        }
    }
}
