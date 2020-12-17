using JMSL.Framework.Unity;

namespace DocumentImport
{
    public class ModuleCatalog : Microsoft.Practices.Prism.Modularity.ModuleCatalog
    {
        public ModuleCatalog()
        {
            AddModuleInternal<Module>();
            AddModuleInternal<JMSL.Framework.DAL.ModuleDAL>();
            AddModuleInternal<JMSL.Framework.Business.ModuleBusiness>();
            AddModuleInternal<ComboxManager.Business.ModuleComboxBusiness>();
        }

        private void AddModuleInternal<T>()
        {
            ((Microsoft.Practices.Prism.Modularity.ModuleCatalog)this).AddModule<T>();
        }
    }
}
