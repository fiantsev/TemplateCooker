using PluginAbstraction;

namespace TemplateCooking.Service.Processing
{
    public interface IInjectionProcessor
    {
        /// <summary>
        /// перерабатывает входной поток инъекций в выходной поток операций
        /// например одна инъекция на вставление таблицы может превратиться в несколько операций (непосредственно вставление таблицы и вставление пустых строк при динамическом количестве строк)
        /// </summary>
        ProcessingStreams Process(IWorkbookAbstraction workbook, ProcessingStreams processingStreams);
    }
}