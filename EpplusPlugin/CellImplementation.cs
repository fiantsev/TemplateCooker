using OfficeOpenXml;
using PluginAbstraction;
using System;

namespace EpplusPlugin
{
    public class CellImplementation : ICellAbstraction
    {
        protected ExcelRange _cell;
        public CellImplementation(ExcelRange cell)
        {
            _cell = cell;
        }


        public int RowIndex => _cell.Start.Row - 1;

        public int ColumnIndex => _cell.Start.Column - 1;

        public bool HasFormula => !string.IsNullOrEmpty(_cell.Formula);

        public CellType Type => CellTypeConverter.GetTypeOf(_cell);

        

        public void SetValue(object value)
        {
            switch (value)
            {
                case int _:
                case long _:
                case double _:
                    var number = Convert.ToDouble(value);
                    _cell.Value = number;
                    //_cell.SetDataType(XLDataType.Number);
                    break;

                case string stringValue:
                    _cell.Value = stringValue;
                    //_cell.SetValue(stringValue);
                    //_cell.SetDataType(XLDataType.Text);
                    break;

                case bool boolValue:
                    _cell.Value = boolValue;
                    //_cell.SetValue(boolValue);
                    //_cell.SetDataType(XLDataType.Boolean);
                    break;

                case DateTime dataTimeValue:
                    _cell.Value = dataTimeValue;
                    //_cell.SetValue(dataTimeValue);
                    //_cell.SetDataType(XLDataType.DateTime);
                    break;

                default:
                    var stringifiedValue = value?.ToString() ?? "";
                    _cell.Value = stringifiedValue;
                    //_cell.SetValue(stringifiedValue);
                    //_cell.SetDataType(XLDataType.Text);
                    break;
            }
        }

        public IMergedCellCollectionAbstraction GetMergedCells(int rowCount, int columnCount, int fixedRowStep, int fixedColumnStep)
        {
            return new MergedCellCollectionImplementation(_cell, rowCount, columnCount, fixedRowStep, fixedColumnStep);
        }

        public string GetStringValue()
        {
            return _cell.GetValue<string>();
        }

        public double GetNumberValue()
        {
            return _cell.GetValue<double>();
        }

        public bool GetBooleanValue()
        {
            return _cell.GetValue<bool>();
        }

        public object GetValue()
        {
            return _cell.Value;
        }

        public void Dispose()
        {
        }

        public void Copy(ExcelRange cell)
        {
            //var innerCell = (cell as CellImplementation)._cell;
            _cell.Copy(cell);
            //_cell.Cop .CopyTo(innerCell);
        }

        public IRangeAbstraction GetMergedRange()
        {
            //_cell.Worksheet.Mer
            var range = _cell.Merge
                ? _cell.Address
                : _cell.Worksheet.Range(_cell, _cell);
            return new RangeImplementation(range);
        }

        public string Address => _cell.Address.ToString();
        public ISheetAbstraction Sheet => new SheetImplementation(_cell.Worksheet);
    }
}
