using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestArrays
{
    [TestClass]
    public class UnitTestIsArray
    {
        [TestMethod]
        public void TestMethod1()
        {
            object[] asd = ArrayFunctions.Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat");
            bool result = ArrayFunctions.IsArray(asd);
            Assert.IsTrue(result);
            
        }
        [TestMethod]
        public void TestMethod2()
        {
            object asd = ArrayFunctions.Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat");
            bool result = ArrayFunctions.IsArray(asd);
            Assert.IsTrue(result);

        }

        [TestMethod]
        public void TestMethod3()
        {
            string asd = "asd";
            bool result = ArrayFunctions.IsArray(asd);
            Assert.IsFalse(result);

        }
    }
}
