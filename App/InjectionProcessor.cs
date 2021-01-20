using PluginAbstraction;
using System;
using System.Linq;
using TemplateCooking.Service.Processing;

namespace XlsxTemplateReporter
{
    public class InjectionProcessor : IInjectionProcessor
    {
        public ProcessingStreams Process(IWorkbookAbstraction workbook, ProcessingStreams processingStreams)
        {
            processingStreams.InjectionStream
                .Select(context => new
                {
                    Type = context.Injection?.GetType().Name,
                })
                .GroupBy(x => x.Type)
                .ToList()
                .ForEach(group => Console.WriteLine($"{group.Key}: {group.Count()}"));

            var result = new DefaultInjectionProcessor().Process(workbook, processingStreams);

            processingStreams.OperationStream
                .Select(operation => new
                {
                    Type = operation?.GetType().FullName,
                })
                .GroupBy(x => x.Type)
                .ToList()
                .ForEach(group => Console.WriteLine($"{group.Key}: {group.Count()}"));

            return result;
        }
    }
}