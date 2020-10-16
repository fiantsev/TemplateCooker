using Abstractions;
using ClosedXML.Excel;
using System;

namespace ClosedXmlPlugin
{
    public class CellImplementation : ICellAbstraction
    {
        private IXLCell _cell;

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
                    var number = (double)value;
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


        public IMergedRowCollectionAbstraction GetMergedRows()
        {
            return new MergedRowCollectionImplementation(_cell);
        }

        public IMergedCellCollectionAbstraction GetMergedCells()
        {
            return new MergedCellCollectionImplementation(_cell);
        }
    }
}