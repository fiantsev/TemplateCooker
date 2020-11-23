using PluginAbstraction;

namespace TemplateCooker.Service.OperationExecutors
{
    public interface IOperationExecutor
    {
        void Execute(IWorkbookAbstraction workbook, object untypedOptions);
    }
}