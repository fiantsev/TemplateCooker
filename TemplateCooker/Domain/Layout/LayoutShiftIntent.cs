using System;
using TemplateCooker.Domain.ResourceObjects;

namespace TemplateCooker.Domain.Layout
{
    public class LayoutShiftIntent
    {
        public LayoutShiftIntent(LayoutShiftOrigin origin, LayoutShiftType type)
        {
            Origin = origin ?? throw new ArgumentNullException(nameof(origin));
            Type = type;
        }

        public LayoutShiftOrigin Origin { get; }
        public LayoutShiftType Type { get; }
    }
}