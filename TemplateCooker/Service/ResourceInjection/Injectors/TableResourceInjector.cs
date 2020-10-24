using System;
using TemplateCooker.Domain.Injections;

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

            var cellCounter = 0;
            foreach (var cell in topLeftCell.GetMergedCells(rowCount, columnCount))
            {
                var rowIndex = cellCounter / columnCount;
                var columnIndex = cellCounter % columnCount;
                cell.SetValue(table[rowIndex][columnIndex]);
                ++cellCounter;
            }
        }
    }
}