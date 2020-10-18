using System.IO;

namespace PluginAbstraction
{
    public interface IPluginAbstraction
    {
        IWorkbookAbstraction CreateEmptyWorkbook();
        IWorkbookAbstraction OpenWorkbook(Stream stream);
    }
}