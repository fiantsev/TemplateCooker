using Abstractions;
using System.Collections.Generic;
using System.Linq;

namespace TemplateCookerUnitTests._Helpers
{
    public class ExcelHelper
    {
        public List<List<object>> ReadCellRangeValues(
            IWorkbookAbstraction workbook,
            (int sheetIndex, int rowIndex, int columnIndex) from,
            (int sheetIndex, int rowIndex, int columnIndex) to
        )
        {
            var result = new List<List<object>>();

            var sheet = workbook
                .GetSheet(from.sheetIndex);

            foreach (var rowIndex in Enumerable.Range(from.rowIndex, to.rowIndex))
            {
                var row = sheet.GetRow(rowIndex);

                result.Add(new List<object>());
                var resultRow = result.Last();

                foreach (var columnIndex in Enumerable.Range(from.columnIndex, to.columnIndex))
                {
                    var cell = row.GetCell(columnIndex);
                    resultRow.Add(cell.GetValue());
                }
            }

            return result;
        }
    }
}
