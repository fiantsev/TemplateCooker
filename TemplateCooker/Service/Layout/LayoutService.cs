using System;
using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.Layout;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker.Service.Layout
{
    public class LayoutService
    {
        class LayoutShift
        {
            public LayoutShiftIntent Intent { get; set; }
            public InjectionContext Item { get; set; }
        }

        /// <summary>
        /// инъекции должны приходить по одному листу
        /// </summary>
        public List<InjectionContext> ProcessLayout(List<InjectionContext> contexts)
        {
            if (contexts.Select(x => x.MarkerRange.StartMarker.Position.SheetIndex).Distinct().Count() > 1)
                throw new Exception("инъекции должны приходить по одному листу");
            var s1 = ExtractLayoutShiftIntents(contexts);
            var s2 = InnerProcessLayout(s1);
            return s2;
        }

        //one-to-one mapping should be preserved
        private List<LayoutShift> ExtractLayoutShiftIntents(List<InjectionContext> contexts)
        {
            return contexts.ToDictionary(context => context, context =>
                {
                    var rowIndex = context.MarkerRange.StartMarker.Position.RowIndex;
                    var columnIndex = context.MarkerRange.StartMarker.Position.ColumnIndex;
                    var rcPosition = new RcPosition(rowIndex, columnIndex);

                    switch (context.Injection)
                    {
                        case TableInjection tableInjection:
                            {
                                var tableRowCount = tableInjection.Resource.Object.Count;
                                var tableColumnCount = tableRowCount == 0 ? 0 : tableInjection.Resource.Object[0].Count;
                                var rcDimensions = new RcDimensions(
                                    Math.Max(1, tableRowCount),
                                    Math.Max(1, tableColumnCount)
                                );
                                return new LayoutShiftIntent(new LayoutElement(rcPosition, rcDimensions), tableInjection.LayoutShift);
                            }
                        default:
                            return new LayoutShiftIntent(new LayoutElement(rcPosition, new RcDimensions(1, 1)), LayoutShiftType.None);
                    }
                })
                .Select(x => new LayoutShift { Intent = x.Value, Item = x.Key })
                .ToList();
        }

        /// <summary>
        /// метод должен вернуть обновленные позиции инъекций и дополнить поток инъекций новыми утилити-инъекциями (напр добавление пустых строк)
        /// </summary>
        /// <param name="layoutMappings"></param>
        /// <returns></returns>
        private List<InjectionContext> InnerProcessLayout(List<LayoutShift> layoutShifts)
        {
            var outputStream = new List<InjectionContext>();

            //HACK
            var sheetIndex = layoutShifts.FirstOrDefault()?.Item.MarkerRange.StartMarker.Position.SheetIndex ?? 0;
            var workbook = layoutShifts.FirstOrDefault()?.Item.Workbook;
            var previousRowShiftAmountAccumulated = 0;

            layoutShifts
                .GroupBy(x => x.Intent.LayoutElement.TopLeft.RowIndex)
                .ToList()
                .ForEach(rowGroup =>
                {
                    var rowIndex = rowGroup.Key;

                    var withRowShiftIntent = rowGroup.Where(x => x.Intent.Type == LayoutShiftType.MoveRows).ToList();

                    var maxAreaHeight = withRowShiftIntent.Select(x => x.Intent.LayoutElement.Area.Height).DefaultIfEmpty(0).Max();
                    //HACK: создаем фейковый маркер (и рендж) только ради передачи rowIndex'a дальше в инжеткор пустых строк
                    var markerRange = new MarkerRange(new Marker("", new SrcPosition(sheetIndex, rowIndex + previousRowShiftAmountAccumulated, 0), MarkerType.Start));

                    var noNeedRowShiftInjection = withRowShiftIntent.Count == 0 || maxAreaHeight == 0;

                    if (noNeedRowShiftInjection)
                    {
                        outputStream.AddRange(rowGroup.Select(x => new InjectionContext { Injection = x.Item.Injection, MarkerRange = x.Item.MarkerRange.WithShift(previousRowShiftAmountAccumulated, 0), Workbook = workbook }));
                    }
                    else
                    {
                        //нужно избавиться здесь от InjectionContext'ov потому что они здесь излишнии
                        outputStream.Add(new InjectionContext { Injection = new EmptyRowsInjection(maxAreaHeight), MarkerRange = markerRange, Workbook = workbook });
                        outputStream.AddRange(rowGroup.Select(x => new InjectionContext { Injection = x.Item.Injection, MarkerRange = x.Item.MarkerRange.WithShift(previousRowShiftAmountAccumulated, 0), Workbook = workbook }));
                        outputStream.Add(new InjectionContext { Injection = new ExtendFormulasDownInjection { SheetIndex = sheetIndex, FromRowIndex = rowIndex + previousRowShiftAmountAccumulated, ToRowIndex = rowIndex + maxAreaHeight + previousRowShiftAmountAccumulated }, MarkerRange = markerRange, Workbook = workbook });
                        previousRowShiftAmountAccumulated += maxAreaHeight;
                    }
                });

            return outputStream;
        }
    }
}