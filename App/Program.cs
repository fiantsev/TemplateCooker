using ClosedXML.Excel;
using PluginAbstraction;
using System;
using System.Data;
using System.IO;
using System.Linq;
using TemplateCooking.Domain.Markers;
using TemplateCooking.Recipes;
using TemplateCooking.Service.Cooker;

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
                //"real-project-report",
                "_current",
                "0503151_fss",
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
            Console.WriteLine($"workbook: {file.In}");
            using var fileStream = File.Open(file.In, FileMode.Open, FileAccess.Read);

            var templateBuilder = new TemplateCooker(fileStream);
            var markerOptions = new MarkerOptions("{{", ".", "}}");

            //при реальном использование есть необходимость извлечь все маркеры прежде чем двигаться дальше
            //маркеры необходимы для того что бы отправить запрос за данными
            var allMarkers = templateBuilder.ExtractMarkers(new ExtractMarkersRecipe.Options { MarkerOptions = markerOptions });
            Console.WriteLine($"Found {allMarkers.Count}: {string.Join(',', allMarkers.Select(x => x.Id))}");

            var injectRecipeOptions = new InjectRecipe.Options
            {
                InjectionProvider = new InjectionProvider(),
                InjectionProcessor = new InjectionProcessor(),
                MarkerOptions = markerOptions,
            };
            templateBuilder.InjectData(injectRecipeOptions);

            var customProperties = new SetCustomPropertiesRecipe.Options
            {
                CustomProperties = new CustomProperties
                {
                    RecalculateFormulasOnSave = false,
                    WorkbookProperties = new WorkbookProperties { ForceFullCalculation = true, FullCalculationOnLoad = true }
                }
            };
            templateBuilder.SetCustomProperties(customProperties);

            var documentStream = templateBuilder.Build();

            using (var outputFileStream = File.Open(file.Out, FileMode.Create, FileAccess.ReadWrite))
                documentStream.CopyTo(outputFileStream);
        }

        static void TreatFile2(InOut file)
        {
            Console.WriteLine($"workbook: {file}");
            using var fileStream = File.Open(file.In, FileMode.Open, FileAccess.Read);

            var workbook = new XLWorkbook(fileStream);
            var range = workbook.Worksheet(1).Range("A1:D2");
            range.CopyTo(workbook.Worksheet(1).Row(3).Cell(1));

            //workbook.CalculationOnSave = true;
            using (var outputFileStream = File.Open(file.Out, FileMode.Create, FileAccess.ReadWrite))
                workbook.SaveAs(outputFileStream, false, false);
        }
    }
}