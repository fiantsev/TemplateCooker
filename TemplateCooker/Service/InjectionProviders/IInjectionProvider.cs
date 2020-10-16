using TemplateCooker.Domain.Injections;

namespace TemplateCooker.Service.InjectionProviders
{
    public interface IInjectionProvider
    {
        Injection Resolve(string key);
    }
}