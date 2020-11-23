using System;
using TemplateCooking.Domain.ResourceObjects;

namespace TemplateCooking.Domain.Layout
{
    public class LayoutShiftIntent
    {
        public LayoutShiftIntent(LayoutElement layoutElement, LayoutShiftType type, RcDimensions firstCellDimensions)
        {
            LayoutElement = layoutElement ?? throw new ArgumentNullException(nameof(layoutElement));
            FirstCellDimensions = firstCellDimensions ?? throw new ArgumentNullException(nameof(firstCellDimensions));
            Type = type;
        }

        public LayoutElement LayoutElement { get; }
        public LayoutShiftType Type { get; }
        //MEGA HACK
        public RcDimensions FirstCellDimensions { get; }
    }
}