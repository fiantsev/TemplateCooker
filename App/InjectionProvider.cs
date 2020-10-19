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
                        var resource = GetResourceObjectStorage()["tableOfIntX3"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.MoveRows };

                    }
                case "table3":
                    {
                        var resource = GetResourceObjectStorage()["tableOfInt"];
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

            var dictionary = new Dictionary<string, ResourceObject>
            {
                { nameof(tableOfInt), new TableResourceObject(tableOfInt) },
                { nameof(tableOfString), new TableResourceObject(tableOfString) },
                { nameof(tableOfObjects), new TableResourceObject(tableOfObjects) },
                { nameof(tableOfIntX3), new TableResourceObject(tableOfIntX3) },
            };


            return dictionary;
        }
    }
}