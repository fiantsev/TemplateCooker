using TemplateCooking.Domain.ResourceObjects;

namespace TemplateCooking.Domain.Injections
{
    public class TableInjection : Injection
    {
        public TableResourceObject Resource { get; set; }

        public LayoutShiftType LayoutShift { get; set; }
        public bool MergeColumnHeaders { get; set; }
        public bool MergeRowHeaders { get; set; }
    }
}