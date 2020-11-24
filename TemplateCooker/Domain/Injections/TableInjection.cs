using TemplateCooking.Domain.ResourceObjects;

namespace TemplateCooking.Domain.Injections
{
    public class TableInjection : Injection
    {
        public TableResourceObject Resource { get; set; }

        public LayoutShiftType LayoutShift { get; set; }
    }
}