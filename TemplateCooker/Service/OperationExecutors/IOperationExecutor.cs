using PluginAbstraction;

namespace TemplateCooking.Service.OperationExecutors
{
    public interface IOperationExecutor
    {
        void Execute(IWorkbookAbstraction workbook, object untypedOptions);
    }
}