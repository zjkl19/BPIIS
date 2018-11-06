using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;
using BPIIS.IRepository;

namespace BPIIS.Repository
{
    public class ProjectRepository : IProjectRepository
    {
        private string result = "";
        private string resultContext = "";
        private MatchCollection Match;
        private Regex regex;
        private Regex regexContext;

        public string GetName(Document doc)
        {
            
            try
            {
                var table0 = doc.GetChildNodes(NodeType.Table, true)[1] as Table;
                Cell cell = table0.Rows[2].Cells[1];
                result = cell.GetText();
                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }


            return result;
        }

        //目前功能暂时同GetName函数
        public string GetBridgeName(Document doc)
        {
            try
            {
                var table0 = doc.GetChildNodes(NodeType.Table, true)[1] as Table;
                Cell cell = table0.Rows[2].Cells[1];
                result = cell.GetText();
                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }
            return result;
        }

        public string GetContractNo(Document doc)
        {
            try
            {
                var table0 = doc.GetChildNodes(NodeType.Table, true)[1] as Table;
                Cell cell = table0.Rows[0].Cells[4];
                result = cell.GetText();
                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }
            return result;
        }

        public bool IsExistRegularPeriod(Document doc)
        {
            string originalWholeText = doc.Range.Text;    //原始全文

            bool existsResult=false;

            regex = new Regex(@"外观检查");
            try
            {
                Match = regex.Matches(originalWholeText);

                result = Match[0].Value;    

                result = SpecialStrReplaceAndTrim(result);         

            }
            catch (Exception)
            {
                result = "未找到";
            }

            if(result=="外观检查")
            {
                existsResult = true;
            }

            return existsResult;
        }

        public bool IsExistStructurePeriod(Document doc)
        {
            string originalWholeText = doc.Range.Text;    //原始全文

            bool existsResult = false;

            regex = new Regex(@"结构定期检测");
            try
            {
                Match = regex.Matches(originalWholeText);

                result = Match[0].Value;

                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }

            if (result == "结构定期检测")
            {
                existsResult = true;
            }

            return existsResult;
        }

        public bool IsExistStaticLoad(Document doc)
        {
            string originalWholeText = doc.Range.Text;    //原始全文

            bool existsResult = false;

            regex = new Regex(@"静动载试验|静力荷载试验|静载试验");
            try
            {
                Match = regex.Matches(originalWholeText);

                result = Match[0].Value;

                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }

            if (result != "未找到")
            {
                existsResult = true;
            }

            return existsResult;
        }

        public bool IsExistDynamicLoad(Document doc)
        {
            string originalWholeText = doc.Range.Text;    //原始全文

            bool existsResult = false;

            regex = new Regex(@"静动载试验|自振特性试验|自振特性");
            try
            {
                Match = regex.Matches(originalWholeText);

                result = Match[0].Value;

                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }

            if (result != "未找到")
            {
                existsResult = true;
            }

            return existsResult;
        }

        public bool IsExistBearingCapacity(Document doc)
        {
            string originalWholeText = doc.Range.Text;    //原始全文

            bool existsResult = false;

            regex = new Regex(@"承载能力检算");
            try
            {
                Match = regex.Matches(originalWholeText);

                result = Match[0].Value;

                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }

            if (result != "未找到")
            {
                existsResult = true;
            }

            return existsResult;
        }

        public bool IsExistRailThrusting(Document doc)
        {
            string originalWholeText = doc.Range.Text;    //原始全文

            bool existsResult = false;

            regex = new Regex(@"栏杆推力|栏杆水平推力");
            try
            {
                Match = regex.Matches(originalWholeText);
                result = Match[0].Value;
                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }

            if (result != "未找到")
            {
                existsResult = true;
            }

            return existsResult;
        }

        private string SpecialStrReplace(string strIn)
        {
            string strOut = null;
            strOut = strIn.Replace("\a", "");
            strOut = strOut.Replace("\r", "");
            return strOut;
        }

        private string SpecialStrReplaceAndTrim(string strIn)
        {
            string strOut = null;
            strOut = SpecialStrReplace(strIn);
            strOut = strOut.Replace(" ", "");
            return strOut;
        }
    }
}
