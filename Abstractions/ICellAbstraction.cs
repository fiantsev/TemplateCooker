using System;

namespace PluginAbstraction
{
    public interface ICellAbstraction : IDisposable
    {
        int RowIndex { get; }
        int ColumnIndex { get; }
        bool HasFormula { get; }

        CellType Type { get; }

        string GetStringValue();
        double GetNumberValue();
        bool GetBooleanValue();
        object GetValue();

        void SetValue(object value);

        IMergedCellCollectionAbstraction GetMergedCells(int rowCount, int columnCount, int fixedRowStep, int fixedColumnStep);

        void Copy(ICellAbstraction cell);
        IRangeAbstraction GetMergedRange();
    }
}