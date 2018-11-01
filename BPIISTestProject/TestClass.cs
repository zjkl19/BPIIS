using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsFormsApp1.Repository;
using Xunit;

namespace WindowsFormsApp1TestProject
{
    public class TestClass
    {
        [Fact]
        public void T1()
        {
            string wt = "sjdsdfkj先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的40000元整先擦的擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的先擦的";
            var c = new ContractRepository();
            var n=c.GetAmount(wt);

            Assert.Equal("40000", n);


        }
    }

}