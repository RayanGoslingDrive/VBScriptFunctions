using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestReplace
    {
        [TestMethod]
        public void TestReplaceFullTextualComp()
        {
            object txt = "This is a beautiful day!";
            object v = StringFunctions.Replace(txt, "t", "##",1,-1,1);
            Assert.AreEqual("##his is a beau##iful day!", v);
        }
        [TestMethod]
        public void TestReplaceFullBinaryComp()
        {
            object txt = "This is a beautiful day!";
            object v = StringFunctions.Replace(txt, "t", "##", 1, -1, 0);
            Assert.AreEqual("This is a beau##iful day!", v);
        }
        [TestMethod]
        public void TestReplaceFullIncorrectComp()
        {
            object txt = "This is a beautiful day!";
            Action act = () => StringFunctions.Replace(txt, "t", "##", 1, -1, 8);
            Assert.ThrowsException<ArgumentException>(act);
        }
        [TestMethod]
        public void TestReplaceRestrictedAmountOfInstances()
        {
            object txt = "This is a beautiful day!";
            object v = StringFunctions.Replace(txt, "i", "##", 1, 2);
            Assert.AreEqual("Th##s ##s a beautiful day!", v);
        }
        [TestMethod]
        public void TestReplaceWithDefaultCountValue()
        {
            object txt = "This is a beautiful day!";
            object v = StringFunctions.Replace(txt, "i", "##", 1, -1);
            Assert.AreEqual("Th##s ##s a beaut##ful day!", v);
        }
        [TestMethod]
        public void TestReplaceWithSmallCountValue()
        {
            object txt = "This is a beautiful day!";
            Action act = () =>StringFunctions.Replace(txt, "i", "##", 1, -200);
            Assert.ThrowsException<ArgumentOutOfRangeException>(act);
        }
        [TestMethod]
        public void TestReplaceStartAt()
        {
            object txt = "This is a beautiful day!";
            object v = StringFunctions.Replace(txt, "i", "##", 15);
            Assert.AreEqual("t##ful day!", v);
        }
        [TestMethod]
        public void TestReplaceStartAtGreater()
        {
            object txt = "This is a beautiful day!";
            object v = StringFunctions.Replace(txt, "i", "##", 124);
            Assert.AreEqual("", v);
        }
        [TestMethod]
        public void TestReplaceStartAtLower()
        {
            object txt = "This is a beautiful day!";

            Action act = () => StringFunctions.Replace(txt, "i", "##", -124);
            Assert.ThrowsException<ArgumentOutOfRangeException>(act);
        }
        [TestMethod]
        public void TestReplaceShort()
        {
            object txt = "This is a beautiful day!";
            object v = StringFunctions.Replace(txt, "i", "##");
            Assert.AreEqual("Th##s ##s a beaut##ful day!", v);
        }
        [TestMethod]
        public void TestReplaceShortNotFound()
        {
            object txt = "This is a beautiful day!";
            object v = StringFunctions.Replace(txt, "z", "##");
            Assert.AreEqual("This is a beautiful day!", v);
        }
 
    }
}
