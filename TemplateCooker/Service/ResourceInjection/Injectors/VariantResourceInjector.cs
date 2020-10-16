using System;
using TemplateCooker.Domain.Injections;

namespace TemplateCooker.Service.ResourceInjection.Injectors
{
    public class VariantResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => context =>
        {
            switch (context.Injection)
            {
                case TableInjection _:
                    new TableResourceInjector().Inject(context);
                    break;
                case ImageInjection _:
                    new ImageResourceInjector().Inject(context);
                    break;
                case TextInjection _:
                    new TextResourceInjector().Inject(context);
                    break;
                default:
                    throw new Exception($"Неизвестный тип объекта экспорта: {context.Injection?.GetType().Name}");
            }
        };
    }
}