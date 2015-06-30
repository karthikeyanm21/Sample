using Microsoft.Practices.Prism.Modularity;
using Microsoft.Practices.Prism.Regions;

namespace APLPX.Modules.StagingDBConfig
{
    public class AdminModule :IModule
    {
        private readonly IRegionViewRegistry regionViewRegistry;

        public AdminModule(IRegionViewRegistry registry)
        {
            this.regionViewRegistry = registry;   
        }

        public void Initialize()
        {
            regionViewRegistry.RegisterViewWithRegion("AdminRegion", typeof(Views.StagingDB));
        }
    }
}
