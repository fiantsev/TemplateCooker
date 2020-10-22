using System;

namespace TemplateCooker.Domain.Layout
{
    public class LayoutShiftOrigin
    {
        public LayoutShiftOrigin(RcPosition positionOfOrigin, RcDimensions areaOfOrigin)
        {
            PositionOfOrigin = positionOfOrigin ?? throw new ArgumentNullException(nameof(positionOfOrigin));
            AreaOfOrigin = areaOfOrigin ?? throw new ArgumentNullException(nameof(areaOfOrigin));
        }

        public RcPosition PositionOfOrigin { get; }
        public RcDimensions AreaOfOrigin { get; }
    }
}