using TemplateCooking.Domain.Markers;

namespace TemplateCooking.Domain.Injections
{
    public class InjectionContext
    {
        public MarkerRange MarkerRange { get; set; }
        public Injection Injection { get; set; }
    }
}