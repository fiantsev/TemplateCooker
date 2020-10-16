using System;

namespace TemplateCooker.Service.ResourceInjection
{
    public interface IResourceInjector
    {
        Action<InjectionContext> Inject { get; }
    }
}