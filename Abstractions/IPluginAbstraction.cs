using System.IO;

namespace Abstractions
{
    public interface IPluginAbstraction
    {
        IWorkbookAbstraction CreateEmptyWorkbook();
        IWorkbookAbstraction OpenWorkbook(Stream stream);
    }
}