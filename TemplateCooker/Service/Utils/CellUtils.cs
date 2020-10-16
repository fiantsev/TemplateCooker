using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using TemplateCooker.Domain.Markers;

namespace TemplateCooker.Service.Utils
{
    public class CellUtils
    {
        public static bool IsMarkedCell(IXLCell cell, MarkerOptions markerOptions)
        {
            if (cell.DataType == XLDataType.Text && !cell.HasFormula)
            {
                var stringCellValue = cell.GetString().Trim();
                if (stringCellValue.Length < (markerOptions.Prefix.Length + markerOptions.Suffix.Length))
                    return false;
                var isPrefixMatch = stringCellValue.Substring(0, markerOptions.Prefix.Length) == markerOptions.Prefix;
                var isSuffixMatch = stringCellValue.Substring(stringCellValue.Length - markerOptions.Suffix.Length, markerOptions.Suffix.Length) == markerOptions.Suffix;
                if (isPrefixMatch && isSuffixMatch)
                    return true;
            }
            return false;
        }

        public static string ExtractMarkerValue(IXLCell cell, MarkerOptions markerOptions)
        {
            var stringCellValue = cell.GetString().Trim();
            return stringCellValue.Substring(markerOptions.Prefix.Length, stringCellValue.Length - (markerOptions.Prefix.Length + markerOptions.Suffix.Length));
        }

        /// <summary>
        /// TODO: переписать реализацию
        /// </summary>
        public static void SetDynamicCellValue(IXLCell cell, object value)
        {
            switch (value)
            {
                case string stringValue:
                    cell.SetValue(stringValue);
                    cell.SetDataType(XLDataType.Text);
                    break;
                case int intValue:
                    cell.SetValue(intValue);
                    cell.SetDataType(XLDataType.Number);
                    break;
                case long longValue:
                    cell.SetValue(longValue);
                    cell.SetDataType(XLDataType.Number);
                    break;
                case double doubleValue:
                    cell.SetValue(doubleValue);
                    cell.SetDataType(XLDataType.Number);
                    break;
                default:
                    throw new Exception($"Неизвестный тип: {value?.GetType().Name}");
            }
        }

        public static IEnumerable<IXLRow> EnumerateMergedRows(IXLCell fromCell)
        {
            return new MergedRowCollection(fromCell);
        }

        public static IEnumerable<IXLCell> EnumerateMergedCells(IXLCell fromCell)
        {
            return new MergedCellCollection(fromCell);
        }
    }
}