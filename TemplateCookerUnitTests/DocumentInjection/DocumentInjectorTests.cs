using ClosedXML.Excel;
using ClosedXmlPlugin;
using FluentAssertions;
using System.Collections.Generic;
using System.IO;
using TemplateCooking.Domain.Injections;
using TemplateCooking.Domain.Markers;
using TemplateCooking.Domain.ResourceObjects;
using TemplateCooking.Recipes;
using TemplateCooking.Service.InjectionProviders;
using TemplateCookerUnitTests._Helpers;
using Xunit;

namespace TemplateCookerUnitTests.DocumentInjection
{
    public class DocumentInjectorTests
    {
        [Fact]
        public void Должен_заменить_маркер_на_таблицу()
        {
            //assign
            var excelHelper = new ExcelHelper();
            var templatePath = "Assets/Templates/one-marker.xlsx";
            using var workbook = new WorkbookImplementation(new XLWorkbook(templatePath));
            var resourceObject = new TableResourceObject(new List<List<object>> {
                new List<object> { 1, 2 },
                new List<object> { 3, 4 },
            });
            var injection = new TableInjection { Resource = resourceObject, LayoutShift = LayoutShiftType.None };
            var documentInjectorOptions = new InjectRecipe.Options
            {
                InjectionProcessor = InjectRecipe.Options.DefaultInjectionProcessor,
                InjectionProvider = new FuncInjectionProvider(_ => injection),
                MarkerOptions = new MarkerOptions("{", ".", "}"),
            };
            var documentInjector = new InjectRecipe(documentInjectorOptions);

            //act
            documentInjector.Cook(workbook);

            //assert
            var values = excelHelper.ReadCellRangeValues(workbook, (0, 0, 0), (1, 1, 1));
            values[0][0].Should().Be(1);
            values[0][1].Should().Be(2);
            values[1][0].Should().Be(3);
            values[1][1].Should().Be(4);
        }

        [Fact]
        public void Должен_вставить_изображение_и_удалить_маркер_из_документа()
        {
            //assign
            var excelHelper = new ExcelHelper();
            using var workbook = new WorkbookImplementation(new XLWorkbook("Assets/Templates/one-marker.xlsx"));
            var imageBytes = File.ReadAllBytes("Assets/Images/checker.png");
            var injection = new ImageInjection { Resource = new ImageResourceObject(imageBytes) };
            var documentInjectorOptions = new InjectRecipe.Options
            {
                InjectionProcessor = InjectRecipe.Options.DefaultInjectionProcessor,
                InjectionProvider = new FuncInjectionProvider(_ => injection),
                MarkerOptions = new MarkerOptions("{", ".", "}"),
            };
            var documentInjector = new InjectRecipe(documentInjectorOptions);

            //act
            documentInjector.Cook(workbook);

            //assert
            var values = excelHelper.ReadCellRangeValues(workbook, (0, 0, 0), (0, 0, 0));
            values[0][0].Should().Be(""); //проверяем что маркер удален из документа
        }

        [Fact]
        public void Должен_заменить_маркер_на_текст()
        {
            //assign
            var excelHelper = new ExcelHelper();
            var templatePath = "Assets/Templates/one-marker.xlsx";
            using var workbook = new WorkbookImplementation(new XLWorkbook(templatePath));
            var injection = new TextInjection { Resource = new TextResourceObject("text") };
            var documentInjectorOptions = new InjectRecipe.Options
            {
                InjectionProcessor = InjectRecipe.Options.DefaultInjectionProcessor,
                InjectionProvider = new FuncInjectionProvider(_ => injection),
                MarkerOptions = new MarkerOptions("{", ".", "}"),
            };
            var documentInjector = new InjectRecipe(documentInjectorOptions);

            //act
            documentInjector.Cook(workbook);

            //assert
            var values = excelHelper.ReadCellRangeValues(workbook, (0, 0, 0), (0, 0, 0));
            values[0][0].Should().Be("text");
        }
    }
}
