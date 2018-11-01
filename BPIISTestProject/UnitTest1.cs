using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WindowsFormsApp1.Repository;

namespace WindowsFormsApp1TestProject
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string wt = "sjdsdfkj先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的40000元整先擦的擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的";
            var c = new ContractRepository();
            var n = c.GetAmount(wt);

            Assert.AreEqual("40000", n);
        }
    }
}
