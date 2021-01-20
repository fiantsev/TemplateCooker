using PluginAbstraction;

namespace TemplateCooking.Recipes
{
    public class SetCustomPropertiesRecipe
    {
        public class Options
        {
            public IWorkbookAbstraction Workbook { get; set; }
            public CustomProperties CustomProperties { get; set; }
        }

        private Options _options;

        public SetCustomPropertiesRecipe(Options options)
        {
            _options = options;
        }

        public void Cook(IWorkbookAbstraction workbook)
        {
            workbook.SetCustomProperties(_options.CustomProperties);
        }
    }
}