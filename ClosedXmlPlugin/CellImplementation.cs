using ClosedXML.Excel;
using PluginAbstraction;
using System;
using System.Diagnostics;

namespace ClosedXmlPlugin
{
    [DebuggerDisplay("{_cell}")]
    public class CellImplementation : ICellAbstraction
    {
        protected IXLCell _cell;

        public int RowIndex => _cell.Address.RowNumber - 1;

        public int ColumnIndex => _cell.Address.ColumnNumber - 1;

        public bool HasFormula => _cell.HasFormula;

        public CellType Type => CellTypeConverter.GetTypeOf(_cell);

        public CellImplementation(IXLCell cell)
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
                    _cell.SetValue(number);
                    _cell.SetDataType(XLDataType.Number);
                    break;

                case string stringValue:
                    _cell.SetValue(stringValue);
                    _cell.SetDataType(XLDataType.Text);
                    break;

                case bool boolValue:
                    _cell.SetValue(boolValue);
                    _cell.SetDataType(XLDataType.Boolean);
                    break;

                case DateTime dataTimeValue:
                    _cell.SetValue(dataTimeValue);
                    _cell.SetDataType(XLDataType.DateTime);
                    break;

                default:
                    var stringifiedValue = value?.ToString() ?? "";
                    _cell.SetValue(stringifiedValue);
                    _cell.SetDataType(XLDataType.Text);
                    break;
            }
        }

        public IMergedCellCollectionAbstraction GetMergedCells(int rowCount, int columnCount, int fixedRowStep, int fixedColumnStep)
        {
            return new MergedCellCollectionImplementation(_cell, rowCount, columnCount, fixedRowStep, fixedColumnStep);
        }

        public string GetStringValue()
        {
            return _cell.GetString();
        }

        public double GetNumberValue()
        {
            return _cell.GetDouble();
        }

        public bool GetBooleanValue()
        {
            return _cell.GetBoolean();
        }

        public object GetValue()
        {
            return _cell.Value;
        }

        public void Dispose()
        {
        }

        public void Copy(ICellAbstraction cell)
        {
            var innerCell = (cell as CellImplementation)._cell;
            _cell.CopyTo(innerCell);
        }

        public IRangeAbstraction GetMergedRange()
        {
            var range = _cell.IsMerged()
                ? _cell.MergedRange()
                : _cell.Worksheet.Range(_cell, _cell);
            return new RangeImplementation(range);
        }
    }
}