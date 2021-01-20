using System;
using TemplateCooking.Domain.Injections;

namespace TemplateCooking.Service.ResourceInjection
{
    public interface IResourceInjector
    {
        Action<InjectionContext> Inject { get; }
    }
}