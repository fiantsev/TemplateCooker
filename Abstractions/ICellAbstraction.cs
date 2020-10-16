namespace Abstractions
{
    public interface ICellAbstraction
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


        //подумать правильная ли это реализация
        IMergedRowCollectionAbstraction GetMergedRows();
        IMergedCellCollectionAbstraction GetMergedCells();
    }
}