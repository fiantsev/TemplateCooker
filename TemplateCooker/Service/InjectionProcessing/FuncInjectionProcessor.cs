using System;
using System.Collections.Generic;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker.Service.InjectionProcessing
{
    public class FuncInjectionProcessor : IInjectionProcessor
    {
        private Func<IEnumerable<InjectionContext>, IEnumerable<InjectionContext>> _process;

        public FuncInjectionProcessor(Func<IEnumerable<InjectionContext>, IEnumerable<InjectionContext>> process)
        {
            _process = process;
        }

        public IEnumerable<InjectionContext> Process(IEnumerable<InjectionContext> injectionContexts)
        {
            return _process.Invoke(injectionContexts);
        }
    }
}