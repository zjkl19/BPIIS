using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPIIS.Model
{
    public class VersionInfo
    {
        //服务器版本号
        public string ServerVersion { get; set; }

        //更新网站
        //http://192.168.12.11:8300/BPIIS.rar
        //或
        //http://218.66.5.89:8300/BPIIS.rar
        public string ServerSite { get; set; }


    }
}
