using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using VBScriptFunctions;

namespace UnitTestArrays
{
    [TestClass]
    public class UnitTestArray
    {
        [TestMethod]
        public void TestMethod1()
        {
            object[] v = ArrayFunctions.Array(1, null, "3", true, 5.7);
            object[] v2 = { 1, null, "3", true, 5.7 };
            CollectionAssert.AreEqual(v2, v);
        }
        [TestMethod]
        public void TestMethod2()
        {
            object[] v = ArrayFunctions.Array();
            object[] v2 = { };
            CollectionAssert.AreEqual(v2, v);
        }
    }
}
