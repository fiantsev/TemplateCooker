using System.Collections.Generic;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker.Service.InjectionProcessing
{
    public interface IInjectionProcessor
    {
        IEnumerable<InjectionContext> Process(IEnumerable<InjectionContext> injectionContexts);
    }
}