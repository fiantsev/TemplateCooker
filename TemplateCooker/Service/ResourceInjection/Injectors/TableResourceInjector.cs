using PluginAbstraction;
using System;
using System.Linq;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.ResourceObjects;

namespace TemplateCooker.Service.ResourceInjection.Injectors
{
    public class TableResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext injectionContext) =>
        {
            InsertTable(injectionContext);
        };

        private void InsertTable(InjectionContext injectionContext)
        {
            var markerPosition = injectionContext.MarkerRange.StartMarker.Position;
            var tableInjection = (injectionContext.Injection as TableInjection);
            var table = tableInjection.Resource.Object;
            var sheet = injectionContext.Workbook.GetSheet(markerPosition.SheetIndex);
            var topLeftCell = sheet.GetRow(markerPosition.RowIndex).GetCell(markerPosition.ColumnIndex);

            var rowCount = table.Count;
            var columnCount = rowCount == 0
                ? 0
                : table[0].Count;

            //удаляем маркер
            if (rowCount == 0 || columnCount == 0)
                topLeftCell.SetValue(string.Empty);

            if (tableInjection.LayoutShift == LayoutShiftType.MoveRows)
                CloneFirstRowBelow(sheet, topLeftCell, rowCount, columnCount);

            var cellCounter = 0;
            foreach (var cell in topLeftCell.GetMergedCells(rowCount, columnCount))
            {
                var rowIndex = cellCounter / columnCount;
                var columnIndex = cellCounter % columnCount;
                cell.SetValue(table[rowIndex][columnIndex]);
                ++cellCounter;
            }
        }

        private void CloneFirstRowBelow(ISheetAbstraction sheet, ICellAbstraction topLeftCell, int rowCount, int columnCount)
        {
            if (rowCount == 0 || columnCount == 0) 
                return;

            var firstRowMergedCells = topLeftCell.GetMergedCells(1, columnCount).ToList();
            var firstCellHeight = firstRowMergedCells.First().GetMergedRange().Height;
            var range = sheet.GetRange(firstRowMergedCells.First(), firstRowMergedCells.Last().GetMergedRange().BottomRightCell());

            for(var i = 1; i < rowCount; ++i)
            {
                var toCell = sheet.GetRow(topLeftCell.RowIndex + i* firstCellHeight).GetCell(topLeftCell.ColumnIndex);
                range.CopyTo(toCell);
            }
        }
    }
}