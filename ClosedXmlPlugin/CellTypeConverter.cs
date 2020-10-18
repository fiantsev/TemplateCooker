using PluginAbstraction;
using ClosedXML.Excel;

namespace ClosedXmlPlugin
{
    public static class CellTypeConverter
    {
        public static CellType GetTypeOf(IXLCell cell)
        {
            switch (cell.DataType)
            {
                case XLDataType.Text: return CellType.String;
                case XLDataType.Number: return CellType.Number;
                case XLDataType.Boolean: return CellType.Boolean;
                default: return CellType.String;
            }
        }
    }
}