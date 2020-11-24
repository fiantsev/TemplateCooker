using TemplateCooking.Domain.Injections;

namespace TemplateCooking.Service.InjectionProviders
{
    public interface IInjectionProvider
    {
        Injection Resolve(string key);
    }
}