﻿using System;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.Utils;

namespace TemplateCooker.Service.ResourceInjection.Injectors
{
    public class TableResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext injectionContext) =>
        {
            ShiftLayout(injectionContext);
            InsertTable(injectionContext);
        };

        private void ShiftLayout(InjectionContext injectionContext)
        {
            var markerRange = injectionContext.MarkerRange;
            var injection = (injectionContext.Injection as TableInjection);
            var table = injection.Resource.Object;

            switch (injection.LayoutShift)
            {
                case LayoutShiftType.None:
                    return;
                case LayoutShiftType.MoveRows:
                    var countOfRowsToInsert = table.Count > 1
                        ? table.Count - 1 //-1 потому что одна ячейка уже есть, та в которой находиться сам маркер
                        : 0;
                    if (countOfRowsToInsert != 0)
                        injectionContext.Workbook.GetSheet(markerRange.StartMarker.Position.SheetIndex)
                            .GetRow(markerRange.EndMarker.Position.RowIndex)
                            .InsertRowsBelow(countOfRowsToInsert);
                    return;
                case LayoutShiftType.MoveCells:
                    throw new Exception("Unsupported");
                default:
                    throw new Exception($"Unhandled case: {nameof(injection.LayoutShift)}={injection.LayoutShift.ToString()}");
            }
        }

        private void InsertTable(InjectionContext injectionContext)
        {
            var markerPosition = injectionContext.MarkerRange.StartMarker.Position;
            var table = (injectionContext.Injection as TableInjection).Resource.Object;
            var sheet = injectionContext.Workbook.GetSheet(markerPosition.SheetIndex);
            var topLeftCell = sheet.GetRow(markerPosition.RowIndex).GetCell(markerPosition.ColumnIndex);

            var rowCount = table.Count;
            var columnCount = rowCount == 0
                ? 0
                : table[0].Count;

            //удаляем маркер
            if (rowCount == 0 || columnCount == 0)
                topLeftCell.SetValue(string.Empty);

            var mergedRowsEnumerator = CellUtils.EnumerateMergedRows(topLeftCell).GetEnumerator();

            table.ForEach(dataRow =>
            {
                mergedRowsEnumerator.MoveNext();
                var excelRow = mergedRowsEnumerator.Current;

                var firstCellOfRow = sheet.GetRow(excelRow.FirstCell().RowIndex).GetCell(topLeftCell.ColumnIndex);
                var mergedCellsEnumerator = CellUtils.EnumerateMergedCells(firstCellOfRow).GetEnumerator();

                dataRow.ForEach(dataValue =>
                {
                    mergedCellsEnumerator.MoveNext();
                    CellUtils.SetDynamicCellValue(mergedCellsEnumerator.Current, dataValue);
                });
            });
        }
    }
}