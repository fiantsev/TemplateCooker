using System;
using TemplateCooker.Domain.ResourceObjects;

namespace TemplateCooker.Domain.Layout
{
    public class LayoutShiftIntent
    {
        public LayoutShiftIntent(LayoutElement layoutElement, LayoutShiftType type)
        {
            LayoutElement = layoutElement ?? throw new ArgumentNullException(nameof(layoutElement));
            Type = type;
        }

        public LayoutElement LayoutElement { get; }
        public LayoutShiftType Type { get; }
    }
}