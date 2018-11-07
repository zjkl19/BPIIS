using BPIIS.IRepository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BPIIS.Repository
{
    public class VersionRepository:IVersionRepository
    {
        public string GetVersion()
        {
            try
            {
                WebClient MyWebClient = new WebClient();
                MyWebClient.Credentials = CredentialCache.DefaultCredentials;//获取或设置用于向Internet资源的请求进行身份验证的网络凭据

                //软件版本信息
                Byte[] pageData = MyWebClient.DownloadData("http://192.168.12.11:8300/BPIISUpdate.txt"); //从指定网站下载数据

                string pageHtml = Encoding.Default.GetString(pageData);  //如果获取网站页面采用的是GB2312，则使用这句            

                //string pageHtml = Encoding.UTF8.GetString(pageData); //如果获取网站页面采用的是UTF-8，则使用这句

                Byte[] ipPageData = MyWebClient.DownloadData("http://ip.tool.chinaz.com/");
                string ipPageHtml = Encoding.UTF8.GetString(ipPageData);

                var result = string.Empty;
                var regex = new Regex(@"(?<=<dd class=""fz24"">)([\s\S]+?)(?=</dd>)");
                try
                {
                    var Match = regex.Matches(ipPageHtml);
                    result = Match[0].Value;

                }
                catch (Exception)
                {
                    result = "未找到";
                }

                //MessageBox.Show(result);

            }
            catch (WebException webEx)
            {
                //MessageBox.Show(webEx.Message.ToString());
            }
            return string.Empty;
        }
    }
}
