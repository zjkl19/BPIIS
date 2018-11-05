using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BPIIS.IRepository;
using BPIIS.Repository;

namespace BPIIS.Infrastructure
{
    public class NinjectDependencyResolver : Ninject.Modules.NinjectModule
    {
        public override void Load()
        {
            Bind<IContractRepository>().To<ContractRepository>();
            Bind<IProjectRepository>().To<ProjectRepository>();
        }
    }
}
