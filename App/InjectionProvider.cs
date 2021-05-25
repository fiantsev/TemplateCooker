using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TemplateCooking.Domain.ResourceObjects;
using TemplateCooking.Domain.Injections;
using TemplateCooking.Service.InjectionProviders;

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
                case "table7":
                    {
                        var resource = GetResourceObjectStorage()["tableWithHeaders"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.MoveRows, MergeColumnHeaders = false, MergeRowHeaders = true };

                    }
                case "table8":
                    {
                        var resource = GetResourceObjectStorage()["tableWithHeaders2"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.MoveRows, MergeColumnHeaders = true, MergeRowHeaders = true };

                    }
                case "table9":
                    {
                        var resource = GetResourceObjectStorage()["tableWithHeaders3"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.MoveRows, MergeColumnHeaders = true, MergeRowHeaders = true };

                    }
                case "table10":
                    {
                        var resource = GetResourceObjectStorage()["tableWithHeaders4"];
                        return new TableInjection { Resource = (TableResourceObject)resource, LayoutShift = LayoutShiftType.None, MergeColumnHeaders = true, MergeRowHeaders = true };

                    }
                default:
                    {
                        var injection = TryParse(markerId);
                        if (injection == null)
                            throw new Exception("По маркеру нет данных");
                        return injection;
                    }
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

            var tableWithHeaders = new TableWithHeaders(
                new List<List<object>>{
                    new List<object> { "2016", "2016", "2017", "2017" },
                    new List<object> { "Q1", "Q1", "Q1", "Q1", },
                    new List<object> { "January", "February", "January", "February" },
                },
                new List<List<object>>{
                    new List<object> { "Russia", "Moscow" },
                    new List<object> { "Russia", "Moscow" },
                    new List<object> { "Japan", "Tokyo" },
                    new List<object> { "Japan", "Tokyo" },
                    new List<object> { "Japan", "Nara" },
                    new List<object> { "USA", "NewYork" },
                },
                new List<List<object>>{
                    new List<object> { 11, 12, 13, 14 },
                    new List<object> { 21, 22, 23, 24 },
                    new List<object> { 31, 32, 33, 34 },
                    new List<object> { 41, 42, 43, 44 },
                    new List<object> { 51, 52, 53, 54 },
                    new List<object> { 61, 62, 63, 64 },
                }
            );

            var tableWithHeaders2 = new TableWithHeaders(
                new List<List<object>>{
                    new List<object> { "Скидка", "Скидка" },
                },
                new List<List<object>>{
                    new List<object> { "Итого" },
                    new List<object> { "Moscow" },
                    new List<object> { "Tokyo" },
                    new List<object> { "Nara" },
                    new List<object> { "NewYork" },
                },
                new List<List<object>>{
                    new List<object> { null, 12 },
                    new List<object> { 21, null },
                    new List<object> { 31, null },
                    new List<object> { 41, null },
                    new List<object> { 51, null },
                    new List<object> { 61, null },
                }
            );

            var tableWithHeaders3 = new TableWithHeaders(
                new List<List<object>>{
                    new List<object> { "Скидка", "Скидка" },
                },
                new List<List<object>>{
                    new List<object> { "Moscow", "2016", "1" },
                    new List<object> { "Moscow", "2016", "2" },
                    new List<object> { "Nara", "2016", "1" },
                    new List<object> { "Nara", "2016", "2" },
                },
                new List<List<object>>{
                    new List<object> { 11, 12 },
                    new List<object> { 21, 22 },
                    new List<object> { 31, 32 },
                    new List<object> { 41, 42 },
                }
            );

            var tableWithHeaders4 = new TableWithHeaders(
                new List<List<object>>{
                    new List<object> { "Прибыль", "Прибыль" },
                },
                new List<List<object>>{
                    new List<object> { "Барнаул", "2016" },
                    new List<object> { "Барнаул", "2017" },
                    new List<object> { "Москва", "2016" },
                    new List<object> { "Москва", "2017" },
                },
                new List<List<object>>{
                    new List<object> { 11, 12 },
                    new List<object> { 21, 22 },
                    new List<object> { 31, 32 },
                    new List<object> { 41, 42 },
                }
            );

            var dictionary = new Dictionary<string, ResourceObject>
            {
                { nameof(tableOfInt), new TableResourceObject(tableOfInt) },
                { nameof(tableOfString), new TableResourceObject(tableOfString) },
                { nameof(tableOfObjects), new TableResourceObject(tableOfObjects) },
                { nameof(tableOfIntX3), new TableResourceObject(tableOfIntX3) },
                { nameof(tableOneColumnOfInt), new TableResourceObject(tableOneColumnOfInt) },
                { nameof(tableOneColumnOf2Int), new TableResourceObject(tableOneColumnOf2Int) },
                { nameof(tableOneColumnOf3Int), new TableResourceObject(tableOneColumnOf3Int) },
                { nameof(tableWithHeaders), new TableResourceObject(tableWithHeaders) },
                { nameof(tableWithHeaders2), new TableResourceObject(tableWithHeaders2) },
                { nameof(tableWithHeaders3), new TableResourceObject(tableWithHeaders3) },
                { nameof(tableWithHeaders4), new TableResourceObject(tableWithHeaders4) },
            };


            return dictionary;
        }

        private Injection TryParse(string markerId)
        {
            if (markerId.First() == '{' || markerId.Last() == '}')
            {
                var markerContent = markerId.Substring(1, markerId.Length - 2);
                var parts = markerContent.Split(',');
                var height = int.Parse(parts[0]);
                var width = int.Parse(parts[1]);
                return CreateTableInjection(height, width, markerContent);
            }
            return null;
        }

        private Injection CreateTableInjection(int rowCount, int columnCount, string name)
        {
            var oneRow = Enumerable.Repeat((object)name, columnCount).ToList();
            var allRows = Enumerable.Repeat(oneRow, rowCount).ToList();
            var injection = new TableInjection { LayoutShift = LayoutShiftType.MoveRows, Resource = new TableResourceObject(allRows) };
            return injection;
        }
    }
}
