using ClosedXmlPlugin;
using PluginAbstraction;
using System.Collections.Generic;
using System.IO;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Recipes.Read;
using TemplateCooker.Recipes.Update;

namespace TemplateCooker.Service.Builders
{
    public class TemplateBuilder
    {
        private IWorkbookAbstraction _workbook;
        private bool _recalculateFormulasOnBuild;
        private bool _forceFullCalculation;
        private bool _fullCalculationOnLoad;

        public TemplateBuilder(Stream workbookStream)
        {
            workbookStream.Position = 0;
            var plugin = new ClosedXmlPluginImplementation();
            _workbook = plugin.OpenWorkbook(workbookStream);
        }

        public List<Marker> ReadMarkers(MarkerOptions markerOptions)
        {
            var markers = new ExtractMarkersRecipe(new ExtractMarkersRecipe.Options
            {
                MarkerOptions = markerOptions,
                Workbook = _workbook
            }).Cook();

            return markers;
        }

        public TemplateBuilder InjectData(InjectRecipe.Options options)
        {
            new InjectRecipe(new InjectRecipe.Options
            {
                Workbook = _workbook,
                MarkerOptions = options.MarkerOptions,
                InjectionProvider = options.InjectionProvider,
                ResourceInjector = options.ResourceInjector,
            }).Cook();

            return this;
        }

        public TemplateBuilder RecalculateFormulasOnBuild(bool recalculateFormulasOnBuild = true)
        {
            _recalculateFormulasOnBuild = recalculateFormulasOnBuild;
            return this;
        }

        public TemplateBuilder SetupFormulaCalculations(bool forceFullCalculation, bool fullCalculationOnLoad)
        {
            _forceFullCalculation = forceFullCalculation;
            _fullCalculationOnLoad = fullCalculationOnLoad;
            return this;
        }

        public MemoryStream Build()
        {
            var resultStream = new MemoryStream();
            var customProperties = new CustomProperties
            {
                WorkbookProperties = new WorkbookProperties
                {
                    ForceFullCalculation = _forceFullCalculation,
                    FullCalculationOnLoad = _fullCalculationOnLoad,
                },
                RecalculateFormulasOnSave = _recalculateFormulasOnBuild
            };
            _workbook.SetCustomProperties(customProperties);
            _workbook.Save(resultStream);
            resultStream.Position = 0;

            //делаем инстанс более не юзабельным
            _workbook = null;

            return resultStream;
        }
    }
}