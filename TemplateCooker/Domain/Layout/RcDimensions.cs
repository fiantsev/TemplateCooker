using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace TemplateCooker.Domain.Layout
{
    [DebuggerDisplay("RcDimensions({Height},{Width})")]
    public class RcDimensions : IEquatable<RcDimensions>
    {
        public int Height { get; }
        public int Width { get; }

        public RcDimensions(int height, int width)
        {
            if (height < 1) throw new ArgumentException($"{nameof(height)} should be greater than 1");
            if (width < 1) throw new ArgumentException($"{nameof(width)} should be greater than 1");

            Height = height;
            Width = width;
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