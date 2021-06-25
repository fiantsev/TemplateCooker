using OfficeOpenXml;
using PluginAbstraction;

namespace EpplusPlugin
{
    public static class CellTypeConverter
    {
        public static CellType GetTypeOf(ExcelRange cell)
        {
            switch (cell.Value)
            {
                case string _: return CellType.String;
                case int _: return CellType.Number;
                case bool _: return CellType.Boolean;
                default: return CellType.String;
            }
        }
    }
}