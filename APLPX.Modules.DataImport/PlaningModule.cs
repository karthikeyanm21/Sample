using Microsoft.Practices.Prism.Modularity;
using Microsoft.Practices.Prism.Regions;

namespace APLPX.Modules.DataImport
{
    public class PlaningModule : IModule
    {
         private readonly IRegionViewRegistry regionViewRegistry;

         public PlaningModule(IRegionViewRegistry registry)
        {
            this.regionViewRegistry = registry;   
        }

        public void Initialize()
        {
            regionViewRegistry.RegisterViewWithRegion("PlaningRegion", typeof(Views.ImportFile));
        }
    }
}
