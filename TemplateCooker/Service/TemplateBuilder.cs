using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Creation;
using TemplateCooker.Service.Extraction;
using TemplateCooker.Service.FormulaCalculation;

namespace TemplateCooker.Service
{
    public class TemplateBuilder
    {
        private XLWorkbook _workbook;
        private bool _recalculateFormulasOnBuild;
        private FormulaCalculationOptions _formulaCalculationOptions;

        public TemplateBuilder(Stream workbookStream)
        {
            workbookStream.Position = 0;
            _workbook = new XLWorkbook(workbookStream);
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

        public MemoryStream Build(bool validate = true)
        {
            var resultStream = new MemoryStream();

            _workbook.ForceFullCalculation = _formulaCalculationOptions.ForceFullCalculation;
            _workbook.FullCalculationOnLoad = _formulaCalculationOptions.FullCalculationOnLoad;

            _workbook.SaveAs(resultStream, validate, _recalculateFormulasOnBuild);
            resultStream.Position = 0;

            //делаем инстанс более не юзабельным
            _workbook = null;

            return resultStream;
        }
    }
}