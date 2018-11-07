using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPIIS.IRepository
{
    public interface IVersionRepository
    {
        //获取版本号
        string GetVersion();
    }
}
