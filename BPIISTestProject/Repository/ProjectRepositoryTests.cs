using Aspose.Words;
using BPIIS.Repository;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace BPIISTestProject.Repository
{
    public class ProjectRepositoryTests
    {
        string rootPath = Directory.GetCurrentDirectory();
        string fileName = "尤溪光林中桥报告-测试.doc";    

        //ToDo:不存在“外观检查”，AssertFalse
        [Fact]
        public void IsExistRegularPeriod_Returnstrue_WhileRegularPeriodExists()
        {
            Document doc = new Document($"{rootPath}\\项目\\{fileName}");
            var p = new ProjectRepository();
            var r = p.IsExistRegularPeriod(doc);
            Assert.True(r);
        }
    }
}
