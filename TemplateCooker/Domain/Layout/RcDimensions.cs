using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace TemplateCooker.Domain.Layout
{
    [DebuggerDisplay("RcDimensions({RowIndex},{ColumnIndex})")]
    public class RcDimensions : IEquatable<RcDimensions>
    {
        public int Width { get; }
        public int Height { get; }

        public RcDimensions(int width, int height)
        {
            if (width < 0) throw new ArgumentException($"{nameof(width)} should be positive number");
            if (height < 0) throw new ArgumentException($"{nameof(height)} should be positive number");

            Width = width;
            Height = height;
        }


        public override bool Equals(object obj)
        {
            return obj is RcDimensions position &&
                   Width == position.Width &&
                   Height == position.Height;
        }

        public bool Equals(RcDimensions other)
        {
            return ((object)this).Equals(other);
        }

        public override int GetHashCode()
        {
            int hashCode = -1326588986;
            hashCode = hashCode * -1521134295 + Width.GetHashCode();
            hashCode = hashCode * -1521134295 + Height.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(RcDimensions left, RcDimensions right)
        {
            return EqualityComparer<RcDimensions>.Default.Equals(left, right);
        }

        public static bool operator !=(RcDimensions left, RcDimensions right)
        {
            return !(left == right);
        }
    }
}