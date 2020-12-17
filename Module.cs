using ComboxManager.Model;
using JMSL.Framework;
using JMSL.Framework.Contract;
using JMSL.Framework.DAL.Entities;
using Microsoft.Practices.Unity;
namespace DocumentImport
{
    public class Module : ModuleBase
    {
        public override void Initialize()
        {
            Container.RegisterType<IFactoryDbContext, ContextProvider>(new ContainerControlledLifetimeManager());
            Container.RegisterType<IClock, Clock>(new ContainerControlledLifetimeManager());
        }
    }
}