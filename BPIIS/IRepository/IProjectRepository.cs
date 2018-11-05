using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPIIS.IRepository
{
    public interface IProjectRepository
    {
        string GetName(Document doc);

        string GetBridgeName(Document doc);

        string GetContractNo(Document doc);

        bool IsExistRegularPeriod(Document doc);

        bool IsExistStructurePeriod(Document doc);

        bool IsExistStaticLoad(Document doc);

        bool IsExistDynamicLoad(Document doc);
    }
}
