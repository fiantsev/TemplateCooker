using System;
using System.Linq;
using TemplateCooking.Service.Processing;
using TemplateCooking.Service.Processing;

namespace XlsxTemplateReporter
{
    public class InjectionProcessor : IInjectionProcessor
    {
        public ProcessingStreams Process(ProcessingStreams processingStreams)
        {
            processingStreams.InjectionStream
                .Select(context => new
                {
                    Type = context.Injection?.GetType().Name,
                })
                .GroupBy(x => x.Type)
                .ToList()
                .ForEach(group => Console.WriteLine($"{group.Key}: {group.Count()}"));

            var result = new DefaultInjectionProcessor().Process(processingStreams);

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