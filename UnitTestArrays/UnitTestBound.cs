using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestArrays
{
    [TestClass]
    public class UnitTestBound
    {
        [TestMethod]
        public void TestMethod1()
        {
            object obj1 = ArrayFunctions.Array("asd", "asd");
            object obj2 = ArrayFunctions.Array("dsa", "dsa");
            object obj = ArrayFunctions.Array(obj1, obj2);
            object asd = ArrayFunctions.LBound(obj, 0);
            Assert.AreEqual(0, asd);
        }
    }
}
