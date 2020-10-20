using System;
using TemplateCooker.Domain.Markers;

namespace TemplateCooker.Domain.LayoutShifts
{
    public class MarkerRangeShifter
    {
        public void Shift(MarkerRange markerRange, LayoutShift layoutShift)
        {
            if (layoutShift is RowLayoutShift rowLayoutShift)
            {
                markerRange.StartMarker.Position.RowIndex += rowLayoutShift.Amount;

                //пока что отключено потому что функциональность ендМаркеров нигде не должна использоваться  - включить в последствии
                //markerRange.EndMarker.Position.RowIndex += rowLayoutShift.ShiftAmout;

                return;
            }

            throw new Exception($"Unsupported type of shifting: {layoutShift?.GetType().Name ?? "null" }");
        }
    }
}