using System;
using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.Layout;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.ResourceInjection;
using TemplateCooker.Service.Utils;

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
            var result = InnerProcessLayout(contexts);
            return result;
        }

        //one-to-one mapping should be preserved
        private List<LayoutShift> ExtractLayoutShiftIntents(List<InjectionContext> contexts)
        {
            return contexts.ToDictionary(context => context, context =>
                {
                    var sheetIndex = context.MarkerRange.StartMarker.Position.SheetIndex;
                    var rowIndex = context.MarkerRange.StartMarker.Position.RowIndex;
                    var columnIndex = context.MarkerRange.StartMarker.Position.ColumnIndex;
                    var rcPosition = new RcPosition(rowIndex, columnIndex);

                    var cellMergedRange = context.Workbook
                        .GetSheet(sheetIndex)
                        .GetRow(rowIndex)
                        .GetCell(columnIndex)
                        .GetMergedRange();

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
                                return new LayoutShiftIntent(new LayoutElement(rcPosition, rcDimensions), tableInjection.LayoutShift, new RcDimensions(cellMergedRange.Height, cellMergedRange.Width));
                            }
                        default:
                            return new LayoutShiftIntent(new LayoutElement(rcPosition, new RcDimensions(1, 1)), LayoutShiftType.None, new RcDimensions(1, 1));
                    }
                })
                .Select(x => new LayoutShift { Intent = x.Value, Item = x.Key })
                .ToList();
        }

        /// <summary>
        /// метод должен вернуть обновленные позиции инъекций и дополнить поток инъекций новыми утилити-инъекциями (напр добавление пустых строк)
        /// </summary>
        private List<InjectionContext> InnerProcessLayout(List<InjectionContext> contexts)
        {
            var outputStream = new List<InjectionContext>();

            var previousRowShiftAmountAccumulated = 0;

            contexts
                .GroupBy(x => x.MarkerRange.StartMarker.Position.RowIndex)
                .ToList()
                .ForEach(contextsOnSameRow =>
                {
                    var firstContext = contextsOnSameRow.First();
                    var workbook = firstContext.Workbook;
                    var sheetIndex = firstContext.MarkerRange.StartMarker.Position.SheetIndex;
                    var rowIndex = firstContext.MarkerRange.StartMarker.Position.RowIndex;

                    var (rowCountToInsert, rowToInsertPosition) = FindHowMuchNewEmptyRowsToInsertAndWhere(contextsOnSameRow.ToList());

                    if (rowCountToInsert > 0)
                    {
                        //нужно избавиться здесь от InjectionContext'ov потому что они здесь излишнии
                        outputStream.Add(new InjectionContext { Injection = new EmptyRowsInjection(rowCountToInsert), MarkerRange = new MarkerRange(new Marker("", rowToInsertPosition.WithShift(previousRowShiftAmountAccumulated, 0), MarkerType.Start)), Workbook = workbook });
                        outputStream.AddRange(contextsOnSameRow.Select(x => new InjectionContext { Injection = x.Injection, MarkerRange = x.MarkerRange.WithShift(previousRowShiftAmountAccumulated, 0), Workbook = workbook }));
                        outputStream.Add(new InjectionContext { Injection = new FillDownFormulasInjection { SheetIndex = sheetIndex, FromRowIndex = rowIndex + previousRowShiftAmountAccumulated, ToRowIndex = rowIndex + rowCountToInsert + previousRowShiftAmountAccumulated }, /*НЕИСПОЛЬЗУЕТСЯ*/MarkerRange = new MarkerRange(new Marker("", new SrcPosition(sheetIndex, 0,0), MarkerType.Start)), Workbook = workbook });
                        previousRowShiftAmountAccumulated += rowCountToInsert;
                    }
                    else
                    {
                        outputStream.AddRange(contextsOnSameRow.Select(x => new InjectionContext { Injection = x.Injection, MarkerRange = x.MarkerRange.WithShift(previousRowShiftAmountAccumulated, 0), Workbook = workbook }));
                    }
                });

            return outputStream;
        }

        private (int, SrcPosition) FindHowMuchNewEmptyRowsToInsertAndWhere(List<InjectionContext> contexts)
        {
            var tables = contexts
                .Where(x =>
                {
                    if (x.Injection is TableInjection tableInjection && tableInjection.LayoutShift == LayoutShiftType.MoveRows)
                        return true;
                    else
                        return false;
                })
                .Select(x => new
                {
                    Context = x,
                    StartMarkerCellHeight = x.Workbook.GetCell(x.MarkerRange.StartMarker.Position).GetMergedRange().Height,
                    RowCount = (x.Injection as TableInjection).Resource.Object.Count
                });

            var max = tables
                .OrderByDescending(x => x.StartMarkerCellHeight * x.RowCount)
                .FirstOrDefault();

            if (max == null)
                return (0, null);

            var countOfNewRowsToInsert = Math.Max(0, max.RowCount - 1) * max.StartMarkerCellHeight;
            var positionToInsertNewRows = max.Context.MarkerRange.StartMarker.Position.WithShift(max.StartMarkerCellHeight - 1 , 0);
            return (countOfNewRowsToInsert, positionToInsertNewRows);
        }
    }
}