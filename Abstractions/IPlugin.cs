using System.IO;

namespace Abstractions
{
    public interface IPlugin
    {
        IWorkbookAbstraction CreateEmptyWorkbook();
        IWorkbookAbstraction OpenWorkbook(Stream stream);
    }
}