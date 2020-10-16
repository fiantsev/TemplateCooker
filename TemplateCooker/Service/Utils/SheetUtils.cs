using ClosedXML.Excel;

namespace TemplateCooker.Service.Utils
{
    public class SheetUtils
    {
        public static int SheetIndex(IXLWorksheet sheet)
        {
            return sheet.Position;
        }
    }
}