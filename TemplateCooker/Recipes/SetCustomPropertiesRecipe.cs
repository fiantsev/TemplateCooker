using PluginAbstraction;

namespace TemplateCooker.Recipes
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

        public void Cook()
        {
            _options.Workbook.SetCustomProperties(_options.CustomProperties);
        }
    }
}