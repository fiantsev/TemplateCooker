using ClosedXML.Excel;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.Markers;

namespace TemplateCooker.Service.ResourceInjection
{
    public class InjectionContext
    {
        public MarkerRange MarkerRange { get; set; }
        public Injection Injection { get; set; }
        public IXLWorkbook Workbook { get; set; }
    }
}