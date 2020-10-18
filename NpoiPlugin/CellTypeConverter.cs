using NPOI.SS.UserModel;
using PluginCellType = PluginAbstraction.CellType;

namespace NpoiPlugin
{
    public static class CellTypeConverter
    {
        public static PluginCellType GetTypeOf(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String: return PluginCellType.String;
                case CellType.Numeric: return PluginCellType.Number;
                case CellType.Boolean: return PluginCellType.Boolean;
                default: return PluginCellType.String;
            }
        }
    }
}