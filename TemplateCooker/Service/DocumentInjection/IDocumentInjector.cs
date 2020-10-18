using PluginAbstraction;

namespace TemplateCooker
{
    public interface IDocumentInjector
    {
        void Inject(IWorkbookAbstraction workbook);
    }
}