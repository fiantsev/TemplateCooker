using System;

namespace TemplateCooker.Service.ResourceInjection.Injectors
{
    public class NoopInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext injectionContext) =>
        {
            //noop
        };
    }
}