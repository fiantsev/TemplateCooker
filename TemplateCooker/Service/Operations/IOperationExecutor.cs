using PluginAbstraction;

namespace TemplateCooker.Service.Operations
{
    public interface IOperationExecutor
    {
        void Execute(IWorkbookAbstraction workbook, object untypedOptions);
    }
}