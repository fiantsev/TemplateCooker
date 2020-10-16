using Abstractions;
using System.IO;

namespace ClosedXmlPlugin
{
    class PluginImplementation : IPlugin
    {
        public IWorkbookAbstraction CreateEmptyWorkbook()
        {
            return new WorkbookImplementation();
        }

        public IWorkbookAbstraction OpenWorkbook(Stream stream)
        {
            return new WorkbookImplementation(stream);
        }
    }
}