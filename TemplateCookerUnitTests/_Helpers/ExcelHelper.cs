using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

namespace TemplateCookerUnitTests._Helpers
{
    public class ExcelHelper
    {
        public List<List<object>> ReadCellRangeValues(
            XLWorkbook workbook,
            (int sheetIndex, int rowIndex, int columnIndex) from,
            (int sheetIndex, int rowIndex, int columnIndex) to
        )
        {
            var result = new List<List<object>>();

            var sheet = workbook
                .Worksheet(from.sheetIndex);

            foreach (var row in sheet.Rows(from.rowIndex, to.rowIndex))
            {
                result.Add(new List<object>());
                var resultRow = result.Last();
                foreach (var cell in row.Cells(from.columnIndex, to.columnIndex))
                    resultRow.Add(cell.Value);
            }

            return result;
        }
    }
}
