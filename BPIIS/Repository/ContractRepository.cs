using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using BPIIS.IRepository;

namespace BPIIS.Repository
{
    public class ContractRepository : IContractRepository
    {
        private string result = "";
        private string resultContext = "";
        private MatchCollection Match;
        private Regex regex;
        private Regex regexContext;

        private string[,] multiArrayA = 
            { { "李鹏", "528" }, 
            { "陈叶", "825" },
            };
        
        public string GetNo(string wholeText)
        {
            regex = new Regex(@"(?<=合同编号：).*?\r");
            try
            {
                Match = regex.Matches(wholeText);

                result = Match[0].Value;    //此处代码易产生异常

                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }


            return result;
        }

        public string GetName(string wholeText)
        {
            //regex = new Regex(@"(?<=项目名称：)([\s\S]*?)(?=委托方（甲方）：)");
            //regex = new Regex(@"(?is)(?<=项目名称：)(([\s\S]*)\n)?");
            regex = new Regex(@"(?<=项目名称：).*?\r|(?<=工程名称：).*?\r");
            try
            {
                Match = regex.Matches(wholeText);

                result = Match[0].Value;    //此处代码易产生异常

                result = SpecialStrReplaceAndTrim(result);
                
            }
            catch (Exception)
            {
                result = "未找到";
            }


            return result;
        }
        public  string GetAmount(string wholeText)
        {
            //壹贰叁肆伍陆柒捌玖拾佰仟
            regex = new Regex(@"[1-9]{1}[0-9]{3,7}(?=.{1,30}元整)");

            //regex = new Regex(@"[\s\S]{30}元整[\s\S]{30}");

            uint maxAmount = 1;    //最大金额

            try
            {
                Match = regex.Matches(wholeText);

                for(int i=0;i<Match.Count;i++)
                {
                    if(Convert.ToUInt32(Match[i].Value)>maxAmount)
                    {
                        maxAmount = Convert.ToUInt32(Match[i].Value);
                    }
                }

                result = maxAmount.ToString();

            }
            catch (Exception)
            {
                result = "未找到";
            }

            return result;

        }

        public string GetProjectLocation(string wholeText)
        {
            regex = new Regex(@"(?<=技术服务地点：)(.*?)\r|(?<=工程地点：)(.*?)\r");

            try
            {
                Match = regex.Matches(wholeText);
                result = Match[0].Value;

                //regex = new Regex(@"[\u4e00-\u9fa5]+");
                //Match = regex.Matches(result);
                //result = Match[0].Value;

                result = SpecialStrReplaceAndTrim(result);
            }
            catch (Exception)
            {
                result = "未找到";
            }

            return result;
        }

        public string GetSignedDate(string wholeText)
        {
            //分离出的年月日

            regex = new Regex(@"(?<=签订时间：)(.*?)\r|(?<=签订日期：)(.*?)\r");
           
            try
            {
                Match = regex.Matches(wholeText);
                result = Match[0].Value;
                result = SpecialStrReplaceAndTrim(result);

                result = result.Replace("年", "/");
                result = result.Replace("月", "/");
                result = result.Replace("日", "");
            }
            catch (Exception)
            {
                result = "未找到";
            }
            return result;
        }

        public string GetJobContent(string wholeText)
        {
            regex = new Regex(@"(?<=技术服务的目标：)(.*?)\r");

            try
            {
                Match = regex.Matches(wholeText);
                result = Match[0].Value;
                result = SpecialStrReplaceAndTrim(result);
            }
            catch (Exception)
            {
                result = "未找到";
            }

            return result;
        }

        public string GetClient(string wholeText)
        {           
            regex = new Regex(@"(?<=委托方（甲方）：)(.*?)\r|(?<=委托人（甲方）：)(.*?)\r");

            try
            {
                Match = regex.Matches(wholeText);
                result = Match[0].Value;
                result = SpecialStrReplaceAndTrim(result);
            }
            catch (Exception)
            {
                result = "未找到";
            }

            return result;
        }

        public string GetClientContactPerson(string wholeText)
        {
            regex = new Regex(@"(?<=委托方（甲方）：([\s\S]*)项目联系人：).*?\r(?=([\s\S]*)受托方（乙方）：)");

            try
            {
                Match = regex.Matches(wholeText);
                result = Match[0].Value;
                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }

            return result;
        }

        public string GetClientContactPersonPhone(string wholeText)
        {
            regex = new Regex(@"(?<=委托方（甲方）：([\s\S]*)联系方式：)([0-9]{11})\r(?=([\s\S]*)受托方（乙方）：)");

            try
            {
                Match = regex.Matches(wholeText);
                result = Match[0].Value;
                result = SpecialStrReplaceAndTrim(result);

            }
            catch (Exception)
            {
                result = "未找到";
            }

            return result;
        }

        public string GetDeadline(string wholeText)
        {
            regex = new Regex(@"(?<=技术服务期限：).*?\r");

            try
            {
                Match = regex.Matches(wholeText);
                result = Match[0].Value;
                result = SpecialStrReplaceAndTrim(result);
            }
            catch (Exception)
            {
                result = "未找到";
            }

            return result;
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
            strOut=strOut.Replace(" ","");
            return strOut;
        }
    }
}
