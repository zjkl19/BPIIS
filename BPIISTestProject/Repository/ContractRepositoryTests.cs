using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BPIIS.Repository;
using Xunit;

namespace BPIISTestProject.Repository
{
    public class ContractRepositoryTests
    {
        //例：莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测
        [Fact]
        public void GetNo_ReturnsCorrectNo_WhileNoExists()
        {        
            string wt = "\r合同编号：HT02CB1800111\r\r技术服务合同\r\r\r\r项目名称：莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测                     \r委托方（甲方）：";
            var c = new ContractRepository();

            var no = c.GetNo(wt);

            Assert.Equal("HT02CB1800111", no);
        }

        #region GetName
        //例：莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测
        //前缀：项目名称
        [Fact]
        public void GetName_ReturnsCorrectName_WhilePrefix01()
        {
            string wt = "\r合同编号：\r\r技术服务合同\r\r\r\r项目名称：莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、" +
                "桥墩外观及裂缝检测                     \r委托方（甲方）：  仙游县城市建设投资有限公司 ";
            var c = new ContractRepository();

            var name = c.GetName(wt);

            Assert.Equal("莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测", name);
        }

        //例：（改）莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测
        //前缀：工程名称
        [Fact]
        public void GetName_ReturnsCorrectName_WhilePrefix02()
        {
            string wt = "\r合同编号：\r\r技术服务合同\r\r\r\r工程名称：莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、" +
                "桥墩外观及裂缝检测                     \r委托方（甲方）：  仙游县城市建设投资有限公司 ";
            var c = new ContractRepository();

            var name = c.GetName(wt);

            Assert.Equal("莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测", name);
        }

        #endregion

        #region GetAmount
        //例：福州市后屿路改造工程K0+489.7桥外观检查、栏杆推力及静动载试验检测项目
        [Fact]
        public void GetAmount_ReturnsCorrectAmount_WhileAmountForwards()
        {
            string wt = "第四条  甲方向乙方支付技术服务报酬方式：\rl．技术服务费总额为：￥113175（人民币壹拾壹万叁仟壹佰柒拾伍元整）" +
                "。详见附件1预算表。\r";
            var c = new ContractRepository();
            var n = c.GetAmount(wt);

            Assert.Equal("113175", n);
        }

        //例：莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测
        [Fact]
        public void GetAmount_ReturnsCorrectAmount_WhileAmountBehinds()
        {
            string wt = "第四条  甲方向乙方支付技术服务报酬及支付方式为：\r1．本项目技术服务费总价为： 人民币：伍万陆仟元整包干（¥56000元）。\r2．技术服务费由甲方 一次性 支付乙方。\r具体支付方式和时间：乙方提交报告后，甲方在30个工作日内一次性支付给乙方。";
            var c = new ContractRepository();
            var n = c.GetAmount(wt);

            Assert.Equal("56000", n);
        }
        #endregion

        //格式：签订时间：2018年7月17日
        //例：福州市后屿路改造工程K0+489.7桥外观检查、栏杆推力及静动载试验检测项目
        [Fact]
        public void GetSignedDate_ReturnsCorrectSignedDate_WhileFormat01()
        {
            string wt = "受托方（乙方）：福建省建筑工程质量检测中心有限公司\r签订时间：2018年7月17日\r签订地点：福州市\r                   \r委托方（甲方）：";
            var c = new ContractRepository();

            var signedDate = c.GetSignedDate(wt);

            Assert.Equal("2018/7/17", signedDate);
        }

        #region GetProjectLocation
        //例：福州市后屿路改造工程K0+489.7桥外观检查、栏杆推力及静动载试验检测项目
        [Fact]
        public void GetProjectLocation_ReturnsCorrectProjectLocation()
        {
            string wt = "第二条  乙方应按下列要求完成技术服务工作：\r1．技术服务地点：福州市\r2．技术服务期限：合同正式签订后60个工作日内";
            var c = new ContractRepository();

            var pl = c.GetProjectLocation(wt);

            Assert.Equal("福州市", pl);
        }

        //例（改）：福州市后屿路改造工程K0+489.7桥外观检查、栏杆推力及静动载试验检测项目
        [Fact]
        public void GetProjectLocation_ReturnsCorrectProjectLocation_WhilePunctuationExists()
        {
            string wt = "第二条  乙方应按下列要求完成技术服务工作：\r1．技术服务地点：福州市。\r2．技术服务期限：合同正式签订后60个工作日内";
            var c = new ContractRepository();

            var pl = c.GetProjectLocation(wt);

            Assert.Equal("福州市", pl);
        }

        //例：莆田市城区人行过街天桥提升工程（莆田三中段）
        [Fact]
        public void GetProjectLocation_ReturnsCorrectProjectLocation_WhileFormat02()
        {
            string wt = "1、工程名称：莆田市城区人行过街天桥提升工程（莆田三中段）\r、" +
                "2、工程地点：莆田市                             \r第二条  工作内容及范围";
            var c = new ContractRepository();

            var pl = c.GetProjectLocation(wt);

            Assert.Equal("莆田市", pl);
        }
        #endregion

        //例：福州市后屿路改造工程K0+489.7桥外观检查、栏杆推力及静动载试验检测项目
        [Fact]
        public void GetJobContent_ReturnsCorrectJobContent()
        {
            string wt = "第一条  甲方委托乙方进行技术服务的内容如下：\r" +
                "l．技术服务的目标：根据相关国家标准、规范，完成福州市后屿路改造工程K0+489.7桥外观检查、栏杆推力及静动载试验。\r" +
                "2．技术服务的范围：根据相关国家标准、规范，对福州市后屿路改造工程K0+489.7桥进行外观检查、栏杆推力及静动载试验。\r" +
                "3．技术服务的内容：根据相关国家标准、规范，对福州市后屿路改造工程K0+489.7桥进行外观检查、栏杆推力及静动载试验。\r";
            var c = new ContractRepository();

            var jc = c.GetJobContent(wt);

            Assert.Equal("根据相关国家标准、规范，完成福州市后屿路改造工程K0+489.7桥外观检查、栏杆推力及静动载试验。", jc);
        }

        #region GetClient
        //格式：委托方（甲方）：xxx\r
        //例：莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测
        [Fact]
        public void GetClient_ReturnsCorrectClient_WhileFormat01()
        {
            string wt = "项目名称：莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测\r" +
                "委托方（甲方）：  仙游县城市建设投资有限公司\r" +
                "受托方（乙方）： 福建省建筑工程质量检测中心有限公司\r";
            var c = new ContractRepository();

            var client = c.GetClient(wt);

            Assert.Equal("仙游县城市建设投资有限公司", client);

        }
        //格式：委托人（甲方）：xxx\r
        //例：（改）莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测
        [Fact]
        public void GetClient_ReturnsCorrectClient_WhileFormat02()
        {
            string wt = "项目名称：莆永高速公路仙溪大桥57#桥墩垂直度检测、桥墩水平位移监测、桥墩外观及裂缝检测\r" +
                "委托人（甲方）：  仙游县城市建设投资有限公司\r" +
                "受托人（乙方）： 福建省建筑工程质量检测中心有限公司\r";
            var c = new ContractRepository();

            var client = c.GetClient(wt);

            Assert.Equal("仙游县城市建设投资有限公司", client);

        }

        #endregion

        //例：福州市后屿路改造工程K0+489.7桥外观检查、栏杆推力及静动载试验检测项目
        [Fact]
        public void GetClientContactPersonPhone_ReturnsCorrectInfo()
        {
            string wt = "法定代表人：林涛\r项目联系人：林继铭\r联系方式：13799302491\r通讯地址：福州市台江区德榜路53号\r";
            var c = new ContractRepository();

            var ph = c.GetClientContactPersonPhone(wt);

            Assert.Equal("13799302491", ph);
        }


    }
}
