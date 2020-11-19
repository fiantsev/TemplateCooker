using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.InjectionProviders;

namespace XlsxTemplateReporter
{
    public class InjectionProvider : IInjectionProvider
    {
        public Injection Resolve(string markerId)
        {
            switch (markerId)
            {
                case "table1":
                    {
                        var resource = GetResourceObjectStorage()["tableOfInt"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.MoveRows };

                    }
                case "table2":
                    {
                        var resource = GetResourceObjectStorage()["tableOfInt"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.MoveRows };

                    }
                case "table3":
                    {
                        var resource = GetResourceObjectStorage()["tableOneColumnOfInt"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.MoveRows };

                    }
                case "table4":
                    {
                        var resource = GetResourceObjectStorage()["tableOneColumnOfInt"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.None };

                    }
                case "table5":
                    {
                        var resource = GetResourceObjectStorage()["tableOneColumnOf2Int"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.MoveRows };

                    }
                case "table6":
                    {
                        var resource = GetResourceObjectStorage()["tableOneColumnOf3Int"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.MoveRows };

                    }
                case "image1":
                    {
                        var imageBytes = File.ReadAllBytes("./Templates/image1.jpg");
                        return new ImageInjection { Resource = new ImageResourceObject(imageBytes) };
                    }
                case "image2":
                    {
                        var imageBytes = File.ReadAllBytes("./Templates/image2_884x2392.png");
                        return new ImageInjection { Resource = new ImageResourceObject(imageBytes) };
                    }
                case "text1":
                    return new TextInjection { Resource = new TextResourceObject("www.google.com") };
                default:
                    throw new Exception("По маркеру нет данных");
            }
        }

        static Dictionary<string, ResourceObject> GetResourceObjectStorage()
        {
            var tableOfInt = new List<List<object>>
            {
                new List<object> { 1, 2, 3 },
                new List<object> { 4, 5, 6 },
                new List<object> { 7, 8, 9 },
            };

            var tableOfIntX3 = tableOfInt.Concat(tableOfInt).Concat(tableOfInt).ToList();

            var tableOfString = new List<List<object>>
            {
                new List<object> { "1", "2", "3" },
                new List<object> { "4", "5", "6" },
                new List<object> { "7", "8", "9" },
            };

            var tableOfObjects = new List<List<object>>
            {
                new List<object> { true, null, "1", 2, 3.0, (long)4, 5.0f },
            };

            var tableOneColumnOfInt = new List<List<object>>
            {
                new List<object> { 1 },
                new List<object> { 2 },
                new List<object> { 3 },
                new List<object> { 4 },
                new List<object> { 5 },
            };

            var tableOneColumnOf2Int = new List<List<object>>
            {
                new List<object> { 1 },
                new List<object> { 2 },
            };

            var tableOneColumnOf3Int = new List<List<object>>
            {
                new List<object> { 1 },
                new List<object> { 2 },
                new List<object> { 3 },
            };

            var dictionary = new Dictionary<string, ResourceObject>
            {
                { nameof(tableOfInt), new TableResourceObject(tableOfInt) },
                { nameof(tableOfString), new TableResourceObject(tableOfString) },
                { nameof(tableOfObjects), new TableResourceObject(tableOfObjects) },
                { nameof(tableOfIntX3), new TableResourceObject(tableOfIntX3) },
                { nameof(tableOneColumnOfInt), new TableResourceObject(tableOneColumnOfInt) },
                { nameof(tableOneColumnOf2Int), new TableResourceObject(tableOneColumnOf2Int) },
                { nameof(tableOneColumnOf3Int), new TableResourceObject(tableOneColumnOf3Int) },
            };


            return dictionary;
        }
    }
}