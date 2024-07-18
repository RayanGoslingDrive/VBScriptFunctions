using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBScriptFunctions;

namespace UnitTestArrays
{
    [TestClass]
    public class UnitTestJoin
    {
        [TestMethod]
        public void TestMethod1()
        {
            object[] a = ArrayFunctions.Array("11421521", 5351714, true, 512512.46, 'a', "Friday", "");
            string astr = ArrayFunctions.Join(a,"-").ToString();
            Assert.AreEqual("11421521-5351714-True-512512,46-a-Friday-", astr);
        }
    }
}
