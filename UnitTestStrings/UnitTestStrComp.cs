using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestStrings
{
    [TestClass]
    public class UnitTestStrComp
    {
        [TestMethod]
        public void StrCompEqual()
        {
            object value = StringFunctions.StrComp("Hello my fellas", "Hello my fellas");
            Assert.AreEqual(0, value);
        }
        [TestMethod]
        public void StrCompSecondInLower()
        {
            object value = StringFunctions.StrComp("Hello my fellas", "hello my fellas");
            Assert.AreEqual(-1, value);
        }
        [TestMethod]
        public void StrCompSecondInLowerIgnoreCase()
        {
            object value = StringFunctions.StrComp("HELLO MY FELLAS", "hello my fellas", 1);
            Assert.AreEqual(0, value);
        }
        [TestMethod]
        public void StrCompFirstIsLesser()
        {
            object value = StringFunctions.StrComp("hi", "Hello my fellas");
            Assert.AreEqual(-1, value);
        }
        [TestMethod]
        public void StrCompNull()
        {
            Action act = () => StringFunctions.StrComp(null, "Hello my fellas");
            Assert.ThrowsException<ArgumentException>(() => act());
        }

        [TestMethod]
        public void StrCompfloat()
        {
            object value =  StringFunctions.StrComp("asd","ASD",0.6);
            Assert.AreEqual(0, value);
        }

        [TestMethod]
        public void StrCompfloat2()
        {
            object value = StringFunctions.StrComp("asd", "ASD", 0.4);
            Assert.AreEqual(1, value);
        }
        [TestMethod]
        public void StrCompCompareString()
        {
            Action act = () => StringFunctions.StrComp("asd", "ASD","one");
            Assert.ThrowsException<FormatException>(() => act());
        }
        [TestMethod]
        public void StrCompCompareStringSucc()
        {
            object value = StringFunctions.StrComp("asd", "ASD", "1");
            Assert.AreEqual(0, value);
        }
        [TestMethod]
        public void StrCompBool()
        {
            Action act = () => StringFunctions.StrComp("asd", "ASD",true);
            Assert.ThrowsException<ArgumentException>(() => act());
        }
    }
}
