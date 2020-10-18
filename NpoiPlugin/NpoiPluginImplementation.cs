using PluginAbstraction;
using System.IO;

namespace NpoiPlugin
{
    public class NpoiPluginImplementation : IPluginAbstraction
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