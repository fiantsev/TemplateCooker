using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.Linq;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Recipes.Update;
using TemplateCooker.Service.Builders;

namespace XlsxTemplateReporter
{
    internal struct InOut
    {
        public string In { get; set; }
        public string Out { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var templates = new[]
            {
                //"marker-cross",
                //"marker-field",
                //"marker-x-cross",
                //"one-marker-merged-cells",
                //"three-markers-on-one-row-one-marker-wo-rows-shift",
                //"two-markers-on-one-row-one-marker-wo-rows-shift",
                //"two-markers-on-one-column",
                //"two-markers-on-one-row",
                //"one-marker",
                "real-project-report",
                "_current",
            };
            var files = templates
                .Select(x => new InOut
                {
                    In = $"./Templates/{x}.xlsx",
                    Out = $"./Output/{x}.out.xlsx"
                })
                .ToList();

            files.ForEach(TreatFile);

            //Console.ReadKey();
        }

        static void TreatFile(InOut file)
        {
            Console.WriteLine($"workbook: {file}");
            using var fileStream = File.Open(file.In, FileMode.Open, FileAccess.Read);

            var templateBuilder = new TemplateBuilder(fileStream);
            var markerOptions = new MarkerOptions("{{", ".", "}}");

            //при реальном использование есть необходимость извлечь все маркеры прежде чем двигаться дальше
            //маркеры необходимы для того что бы отправить запрос за данными
            var allMarkers = templateBuilder.ReadMarkers(markerOptions);
            Console.WriteLine($"Found {allMarkers.Count}: {string.Join(',', allMarkers.Select(x => x.Id))}");

            var resourceInjector = new ResourceInjector();
            var injectionProvider = new InjectionProvider();
            var injectRecipeOptions = new InjectRecipe.Options
            {
                ResourceInjector = resourceInjector,
                InjectionProvider = injectionProvider,
                MarkerOptions = markerOptions,
            };

            var documentStream = templateBuilder
                .InjectData(injectRecipeOptions)
                .SetupFormulaCalculations(forceFullCalculation: true, fullCalculationOnLoad: true)
                .RecalculateFormulasOnBuild(false)
                .Build();

            using (var outputFileStream = File.Open(file.Out, FileMode.Create, FileAccess.ReadWrite))
                documentStream.CopyTo(outputFileStream);
        }

        static void TreatFile2(InOut file)
        {
            Console.WriteLine($"workbook: {file}");
            using var fileStream = File.Open(file.In, FileMode.Open, FileAccess.Read);

            var workbook = new XLWorkbook(fileStream);
            //workbook.Worksheet(1).Row(2).InsertRowsAbove(2);
            workbook.Worksheet(1).Row(2).Cell(2).CopyTo(workbook.Worksheet(1).Row(3).Cell(2));
            //var templateBuilder = new TemplateBuilder(fileStream);
            //var markerOptions = new MarkerOptions("{{", ".", "}}");

            ////при реальном использование есть необходимость извлечь все маркеры прежде чем двигаться дальше
            ////маркеры необходимы для того что бы отправить запрос за данными
            //var allMarkers = templateBuilder.ReadMarkers(markerOptions);
            //Console.WriteLine($"Found {allMarkers.Count}: {string.Join(',', allMarkers.Select(x => x.Id))}");

            //var resourceInjector = new ResourceInjector();
            //var injectionProvider = new InjectionProvider();
            //var documentInjectorOptions = new DocumentInjectorOptions
            //{
            //    ResourceInjector = resourceInjector,
            //    InjectionProvider = injectionProvider,
            //    MarkerOptions = markerOptions,
            //};

            //var documentStream = templateBuilder
            //    .InjectData(documentInjectorOptions)
            //    .SetupFormulaCalculations(forceFullCalculation: true, fullCalculationOnLoad: true)
            //    .RecalculateFormulasOnBuild(false)
            //    .Build();

            //workbook.CalculationOnSave = true;
            using (var outputFileStream = File.Open(file.Out, FileMode.Create, FileAccess.ReadWrite))
                workbook.SaveAs(outputFileStream, false, false);
        }
    }
}