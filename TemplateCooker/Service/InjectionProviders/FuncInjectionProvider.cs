using System;
using TemplateCooking.Domain.Injections;

namespace TemplateCooking.Service.InjectionProviders
{
    public class FuncInjectionProvider : IInjectionProvider
    {
        private Func<string, Injection> _resolver;

        public FuncInjectionProvider(Func<string, Injection> resolver)
        {
            _resolver = resolver;
        }

        public Injection Resolve(string key)
        {
            return _resolver.Invoke(key);
        }
    }
}