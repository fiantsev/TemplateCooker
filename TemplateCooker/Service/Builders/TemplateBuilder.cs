using PluginAbstraction;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Creation;
using TemplateCooker.Service.Extraction;

namespace TemplateCooker.Service.Builders
{
    public class TemplateBuilder
    {
        private IWorkbookAbstraction _workbook;
        private bool _recalculateFormulasOnBuild;
        private FormulaCalculationOptions _formulaCalculationOptions;

        public TemplateBuilder(Stream workbookStream, IPluginAbstraction plugin)
        {
            workbookStream.Position = 0;
            _workbook = plugin.OpenWorkbook(workbookStream);
            _formulaCalculationOptions = new FormulaCalculationOptions();
        }

        public List<Marker> ReadMarkers(MarkerOptions markerOptions)
        {
            var markerExtractor = new MarkerExtractor(_workbook, markerOptions);
            return markerExtractor.GetMarkers().ToList();
        }

        public TemplateBuilder InjectData(DocumentInjectorOptions options)
        {
            var documentInjector = new DocumentInjector(options);
            documentInjector.Inject(_workbook);

            return this;
        }

        public TemplateBuilder RecalculateFormulasOnBuild(bool recalculateFormulasOnBuild = true)
        {
            _recalculateFormulasOnBuild = recalculateFormulasOnBuild;
            return this;
        }

        public TemplateBuilder SetupFormulaCalculations(FormulaCalculationOptions formulaCalculationOptions)
        {
            _formulaCalculationOptions = formulaCalculationOptions;
            return this;
        }

        public MemoryStream Build()
        {
            var resultStream = new MemoryStream();
            _workbook.SetProperties(_formulaCalculationOptions.ForceFullCalculation, _formulaCalculationOptions.FullCalculationOnLoad, _recalculateFormulasOnBuild);
            _workbook.Save(resultStream);
            resultStream.Position = 0;

            //делаем инстанс более не юзабельным
            _workbook = null;

            return resultStream;
        }
    }
}