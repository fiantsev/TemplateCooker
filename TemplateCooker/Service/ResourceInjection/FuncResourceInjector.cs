using System;

namespace TemplateCooker.Service.ResourceInjection
{
    public class FuncResourceInjector : IResourceInjector
    {
        public FuncResourceInjector(Action<InjectionContext> inject)
        {
            Inject = inject;
        }

        public Action<InjectionContext> Inject { get; set; }
    }
}