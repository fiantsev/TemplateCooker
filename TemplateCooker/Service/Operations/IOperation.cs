using PluginAbstraction;

namespace TemplateCooker.Service.Operations
{
    public interface IOperation
    {
        void Execute(IWorkbookAbstraction workbook, object untypedOptions);
    }
}