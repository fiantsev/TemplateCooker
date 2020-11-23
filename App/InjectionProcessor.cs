using System;
using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Service.Operations;
using TemplateCooker.Service.Processing;
using TemplateCooker.Service.ResourceInjection;

namespace XlsxTemplateReporter
{
    public class InjectionProcessor : IInjectionProcessor
    {
        public void Process(List<InjectionContext> injectionStream, List<AbstractOperation> operationStream)
        {
            injectionStream
                .Select(context => new
                {
                    Type = context.Injection?.GetType().Name,
                })
                .GroupBy(x => x.Type)
                .ToList()
                .ForEach(group => Console.WriteLine($"{group.Key}: {group.Count()}"));

            new DefaultInjectionProcessor().Process(injectionStream, operationStream);

            operationStream
                .Select(operation => new
                {
                    Type = operation?.GetType().FullName,
                })
                .GroupBy(x => x.Type)
                .ToList()
                .ForEach(group => Console.WriteLine($"{group.Key}: {group.Count()}"));
        }
    }
}