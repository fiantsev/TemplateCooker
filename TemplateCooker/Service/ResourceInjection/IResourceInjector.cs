using System;

namespace TemplateCooking.Service.ResourceInjection
{
    public interface IResourceInjector
    {
        Action<InjectionContext> Inject { get; }
    }
}