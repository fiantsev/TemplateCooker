using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace TemplateCooking.Domain.Layout
{
    [DebuggerDisplay("SrPosition({SheetIndex},{RowIndex})")]
    public class SrPosition : IEquatable<SrPosition>
    {
        public int SheetIndex { get; }
        public int RowIndex { get; }

        public SrPosition(int sheetIndex, int rowIndex)
        {
            if (sheetIndex < 0) throw new ArgumentException($"{nameof(sheetIndex)} should be positive number");
            if (rowIndex < 0) throw new ArgumentException($"{nameof(rowIndex)} should be positive number");

            SheetIndex = sheetIndex;
            RowIndex = rowIndex;
        }

        public SrPosition Clone() => new SrPosition(SheetIndex, RowIndex);
        public SrPosition WithShift(int rowShift) => new SrPosition(SheetIndex, RowIndex + rowShift);

        public override bool Equals(object obj)
        {
            return obj is SrPosition position &&
                   SheetIndex == position.SheetIndex &&
                   RowIndex == position.RowIndex;
        }

        public bool Equals(SrPosition other)
        {
            return ((object)this).Equals(other);
        }

        public override int GetHashCode()
        {
            int hashCode = 1713590872;
            hashCode = hashCode * -1521134295 + SheetIndex.GetHashCode();
            hashCode = hashCode * -1521134295 + RowIndex.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(SrPosition left, SrPosition right)
        {
            return EqualityComparer<SrPosition>.Default.Equals(left, right);
        }

        public static bool operator !=(SrPosition left, SrPosition right)
        {
            return !(left == right);
        }
    }
}