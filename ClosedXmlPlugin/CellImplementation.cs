﻿using Abstractions;
using ClosedXML.Excel;
using System;

namespace ClosedXmlPlugin
{
    public class CellImplementation : ICellAbstraction
    {
        private IXLCell _cell;

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


        public IMergedRowCollectionAbstraction GetMergedRows()
        {
            return new MergedRowCollectionImplementation(_cell);
        }

        public IMergedCellCollectionAbstraction GetMergedCells()
        {
            return new MergedCellCollectionImplementation(_cell);
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
    }
}