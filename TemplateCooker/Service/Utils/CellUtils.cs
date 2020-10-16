using Abstractions;
using System.Collections.Generic;
using TemplateCooker.Domain.Markers;

namespace TemplateCooker.Service.Utils
{
    public class CellUtils
    {
        public static bool IsMarkedCell(ICellAbstraction cell, MarkerOptions markerOptions)
        {
            if (cell.Type == CellType.String && !cell.HasFormula)
            {
                var stringCellValue = cell.GetStringValue().Trim();
                if (stringCellValue.Length < (markerOptions.Prefix.Length + markerOptions.Suffix.Length))
                    return false;
                var isPrefixMatch = stringCellValue.Substring(0, markerOptions.Prefix.Length) == markerOptions.Prefix;
                var isSuffixMatch = stringCellValue.Substring(stringCellValue.Length - markerOptions.Suffix.Length, markerOptions.Suffix.Length) == markerOptions.Suffix;
                if (isPrefixMatch && isSuffixMatch)
                    return true;
            }
            return false;
        }

        public static string ExtractMarkerValue(ICellAbstraction cell, MarkerOptions markerOptions)
        {
            var stringCellValue = cell.GetStringValue().Trim();
            return stringCellValue.Substring(markerOptions.Prefix.Length, stringCellValue.Length - (markerOptions.Prefix.Length + markerOptions.Suffix.Length));
        }

        /// <summary>
        /// TODO: переписать реализацию
        /// </summary>
        public static void SetDynamicCellValue(ICellAbstraction cell, object value)
        {
            cell.SetValue(value);
        }

        public static IEnumerable<IRowAbstraction> EnumerateMergedRows(ICellAbstraction fromCell)
        {
            return fromCell.GetMergedRows();
        }

        public static IEnumerable<ICellAbstraction> EnumerateMergedCells(ICellAbstraction fromCell)
        {
            return fromCell.GetMergedCells();
        }
    }
}