using ClosedXML.Excel;

namespace TemplateCooker
{
    public interface IDocumentInjector
    {
        void Inject(IXLWorkbook workbook);
    }
}