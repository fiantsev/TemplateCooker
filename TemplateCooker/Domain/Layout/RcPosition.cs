using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace TemplateCooker.Domain.Layout
{
    [DebuggerDisplay("RcPosition({RowIndex},{ColumnIndex})")]
    public class RcPosition : IEquatable<RcPosition>
    {
        public int RowIndex { get; }
        public int ColumnIndex { get; }

        public RcPosition(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0) throw new ArgumentException($"{nameof(rowIndex)} should be positive number");
            if (columnIndex < 0) throw new ArgumentException($"{nameof(columnIndex)} should be positive number");

            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }


        public override bool Equals(object obj)
        {
            return obj is RcPosition position &&
                   RowIndex == position.RowIndex &&
                   ColumnIndex == position.ColumnIndex;
        }

        public bool Equals(RcPosition other)
        {
            return ((object)this).Equals(other);
        }

        public override int GetHashCode()
        {
            int hashCode = -1326588986;
            hashCode = hashCode * -1521134295 + RowIndex.GetHashCode();
            hashCode = hashCode * -1521134295 + ColumnIndex.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(RcPosition left, RcPosition right)
        {
            return EqualityComparer<RcPosition>.Default.Equals(left, right);
        }

        public static bool operator !=(RcPosition left, RcPosition right)
        {
            return !(left == right);
        }
    }
}