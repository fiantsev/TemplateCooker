using PluginAbstraction;
using TemplateCooking.Domain.Layout;
using TemplateCooking.Domain.Markers;

namespace TemplateCooking.Service.Utils
{
    public static class CellUtils
    {
        public static string ExtractMarkerValueOrNull(ICellAbstraction cell, MarkerOptions markerOptions)
        {
            if (cell.Type == CellType.String && !cell.HasFormula)
            {
                var stringCellValue = cell.GetStringValue().Trim();
                if (stringCellValue.Length < (markerOptions.Prefix.Length + markerOptions.Suffix.Length))
                    return null;
                var isPrefixMatch = stringCellValue.Substring(0, markerOptions.Prefix.Length) == markerOptions.Prefix;
                var isSuffixMatch = stringCellValue.Substring(stringCellValue.Length - markerOptions.Suffix.Length, markerOptions.Suffix.Length) == markerOptions.Suffix;
                if (isPrefixMatch && isSuffixMatch)
                    return stringCellValue.Substring(markerOptions.Prefix.Length, stringCellValue.Length - (markerOptions.Prefix.Length + markerOptions.Suffix.Length)); ;
            }
            return null;
        }

        public static ICellAbstraction GetCell(this IWorkbookAbstraction workbook, SrcPosition srcPosition)
        {
            return workbook.GetSheet(srcPosition.SheetIndex).GetRow(srcPosition.RowIndex).GetCell(srcPosition.ColumnIndex);
        }
    }
}