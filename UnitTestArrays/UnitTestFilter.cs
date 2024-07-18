using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestArrays
{
    [TestClass]
    public class UnitTestFilter
    {
        [TestMethod]
        public void TestFullInclude()
        {
            object[] a = ArrayFunctions.Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
            object[] b = ArrayFunctions.Filter(a, "S",true,0);
            object[] expected = ArrayFunctions.Array("Sunday", "Saturday");
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestFullNotInclude()
        {
            object[] a = ArrayFunctions.Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
            object[] b = ArrayFunctions.Filter(a, "S", false, 0);
            object[] expected = ArrayFunctions.Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday");
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestFullNotIncludeCaseSensetive()
        {
            object[] a = ArrayFunctions.Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
            object[] b = ArrayFunctions.Filter(a, "S", false, 0);
            object[] expected = ArrayFunctions.Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday");
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestFullIncludeAsInt()
        {
            object[] a = ArrayFunctions.Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
            object[] b = ArrayFunctions.Filter(a, "S", 1, 0);
            object[] expected = ArrayFunctions.Array("Sunday", "Saturday");
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestFullNotIncludeCaseInSensetive()
        {
            object[] a = ArrayFunctions.Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
            object[] b = ArrayFunctions.Filter(a, "S", false, 1);
            object[] expected = ArrayFunctions.Array("Monday", "Friday");
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestFullGibberish()
        {
            object[] a = ArrayFunctions.Array("11421521", 5351714, true, 512512.46, 'a', "Friday", "");
            object[] b = ArrayFunctions.Filter(a, 1, 1, 1);
            object[] expected = ArrayFunctions.Array("11421521", 5351714, 512512.46);
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestFullGibberish2()
        {
            object[] a = ArrayFunctions.Array("11421521", 5351714, true, 512512.46, 'a', "Friday", "");
            object[] b = ArrayFunctions.Filter(a, 1, 67, 1);
            object[] expected = ArrayFunctions.Array("11421521", 5351714, 512512.46);
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestFullGibberish3()
        {
            object[] a = ArrayFunctions.Array("11421521", 5351714, true, 512512.46, 'a', "Friday", "");
            object[] b = ArrayFunctions.Filter(a, 1, 0.1, 1);
            object[] expected = ArrayFunctions.Array("11421521", 5351714, 512512.46);
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestFullGibberish4()
        {
            object[] a = ArrayFunctions.Array("11421521", 5351714, true, 512512.46, 'a', "Friday", "");
            object[] b = ArrayFunctions.Filter(a, 1, "1", 1);
            object[] expected = ArrayFunctions.Array("11421521", 5351714, 512512.46);
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestFullGibberish5()
        {
            object[] a = ArrayFunctions.Array("11421521", 5351714, true, 512512.46, 'a', "Friday", "");
            Action act = () =>  ArrayFunctions.Filter(a, 1, "asdasd", 1);
            Assert.ThrowsException<FormatException>(act);
        }
        [TestMethod]
        public void TestFullGibberish6()
        {
            object[] a = ArrayFunctions.Array("11421521", 5351714, true, 512512.46, 'a', "Friday", "");
            Action act = () => ArrayFunctions.Filter(a, "1", "1", "1");
            Assert.ThrowsException<FormatException>(act);
        }
        [TestMethod]
        public void TestThreeparams()
        {
            object[] a = ArrayFunctions.Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
            object[] b = ArrayFunctions.Filter(a, "S", false);
            object[] expected = ArrayFunctions.Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday");
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestThreeparams2()
        {
            object[] a = ArrayFunctions.Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
            object[] b = ArrayFunctions.Filter(a, "S", true);
            object[] expected = ArrayFunctions.Array("Sunday", "Saturday");
            CollectionAssert.AreEqual(expected, b);
        }
        [TestMethod]
        public void TestTwoParams()
        {
            object[] a = ArrayFunctions.Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
            object[] b = ArrayFunctions.Filter(a, "S");
            object[] expected = ArrayFunctions.Array("Sunday", "Saturday");
            CollectionAssert.AreEqual(expected, b);
        }
    }
}
