using NPOI.SS.UserModel;
using PluginAbstraction;
using System;
using System.Diagnostics;

namespace NpoiPlugin
{
    [DebuggerDisplay("{_cell}")]
    public class CellImplementation : ICellAbstraction
    {
        private ICell _cell;

        public int RowIndex => _cell.RowIndex;

        public int ColumnIndex => _cell.ColumnIndex;

        public bool HasFormula => _cell.CellType == NPOI.SS.UserModel.CellType.Formula;

        public PluginAbstraction.CellType Type => CellTypeConverter.GetTypeOf(_cell);

        public CellImplementation(ICell cell)
        {
            _cell = cell;
        }

        public void SetValue(object value)
        {
            switch (value)
            {
                case int _:
                case long _:
                case double _:
                    var number = Convert.ToDouble(value);
                    _cell.SetCellValue(number);
                    break;

                case string stringValue:
                    _cell.SetCellValue(stringValue);
                    break;

                case bool boolValue:
                    _cell.SetCellValue(boolValue);
                    break;

                case DateTime dataTimeValue:
                    _cell.SetCellValue(dataTimeValue);
                    break;

                default:
                    var stringifiedValue = value?.ToString() ?? string.Empty;
                    _cell.SetCellValue(stringifiedValue);
                    break;
            }
        }

        public IMergedCellCollectionAbstraction GetMergedCells(int rowCount, int columnCount)
        {
            return new MergedCellCollectionImplementation(_cell, rowCount, columnCount);
        }

        public string GetStringValue()
        {
            return _cell.ToString();
        }

        public double GetNumberValue()
        {
            return _cell.NumericCellValue;
        }

        public bool GetBooleanValue()
        {
            return _cell.BooleanCellValue;
        }

        public object GetValue()
        {
            switch (_cell.CellType)
            {
                case NPOI.SS.UserModel.CellType.String: return _cell.StringCellValue;
                case NPOI.SS.UserModel.CellType.Numeric: return _cell.NumericCellValue;
                case NPOI.SS.UserModel.CellType.Boolean: return _cell.BooleanCellValue;
                default: return _cell.StringCellValue;
            }
        }

        public void Dispose()
        {
        }
    }
}