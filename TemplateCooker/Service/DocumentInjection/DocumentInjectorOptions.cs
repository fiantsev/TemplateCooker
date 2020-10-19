using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.InjectionProviders;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker.Service.Creation
{
    public class DocumentInjectorOptions
    {
        public IResourceInjector ResourceInjector { get; set; }
        public IInjectionProvider InjectionProvider { get; set; }
        public MarkerOptions MarkerOptions { get; set; }
    }
}