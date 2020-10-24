using ClosedXmlPlugin;
using PluginAbstraction;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TemplateCooker.Recipes.Update;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker.Recipes
{
    public class CreateWorkbookRecipe
    {
        public class Options
        {
            public IEnumerable<(string sheetName, IEnumerable<InjectionContext> injections)> SheetGenerator { get; set; }
        }

        private Options _options;

        public CreateWorkbookRecipe(Options options)
        {
            _options = options;
        }

        public void Cook()
        {
        }
    }
}