using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace TemplateCooker.Domain.Layout
{
    [DebuggerDisplay("SrcPosition({SheetIndex},{RowIndex},{ColumnIndex})")]
    public class SrcPosition : IEquatable<SrcPosition>
    {
        public int SheetIndex { get; }
        public int RowIndex { get; }
        public int ColumnIndex { get; }

        public SrcPosition(int sheetIndex, int rowIndex, int columnIndex)
        {
            if (sheetIndex < 0) throw new ArgumentException($"{nameof(sheetIndex)} should be positive number");
            if (rowIndex < 0) throw new ArgumentException($"{nameof(rowIndex)} should be positive number");
            if (columnIndex < 0) throw new ArgumentException($"{nameof(columnIndex)} should be positive number");

            SheetIndex = sheetIndex;
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        public SrcPosition Clone() => new SrcPosition(SheetIndex, RowIndex, ColumnIndex);
        public SrcPosition WithShift(int rowShift = 0, int columnShift = 0) => new SrcPosition(SheetIndex, RowIndex + rowShift, ColumnIndex + columnShift);
        public SrPosition ToSrPosition() => new SrPosition(SheetIndex, RowIndex);
        public RcPosition ToRcPosition() => new RcPosition(RowIndex, ColumnIndex);

        public override bool Equals(object obj)
        {
            return obj is SrcPosition position &&
                   SheetIndex == position.SheetIndex &&
                   RowIndex == position.RowIndex &&
                   ColumnIndex == position.ColumnIndex;
        }

        public bool Equals(SrcPosition other)
        {
            return ((object)this).Equals(other);
        }

        public override int GetHashCode()
        {
            int hashCode = 1713590872;
            hashCode = hashCode * -1521134295 + SheetIndex.GetHashCode();
            hashCode = hashCode * -1521134295 + RowIndex.GetHashCode();
            hashCode = hashCode * -1521134295 + ColumnIndex.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(SrcPosition left, SrcPosition right)
        {
            return EqualityComparer<SrcPosition>.Default.Equals(left, right);
        }

        public static bool operator !=(SrcPosition left, SrcPosition right)
        {
            return !(left == right);
        }
    }
}