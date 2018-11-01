using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WindowsFormsApp1.IRepository;
using WindowsFormsApp1.Repository;

namespace WindowsFormsApp1.Infrastructure
{
    public class NinjectDependencyResolver : Ninject.Modules.NinjectModule
    {
        public override void Load()
        {
            Bind<IContractRepository>().To<ContractRepository>();

        }
    }
}
