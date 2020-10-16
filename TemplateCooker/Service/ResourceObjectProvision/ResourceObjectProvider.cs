using System;
using TemplateCooker.Domain.ResourceObjects;

namespace TemplateCooker.Service.ResourceObjectProvision
{
    //Ответственность: по Marker.Id вернуть типизированный ResourceObject
    public class ResourceObjectProvider : IResourceObjectProvider
    {
        public Func<string, ResourceObject> Resolve { get; }

        public ResourceObjectProvider(Func<string, ResourceObject> resolve)
        {
            Resolve = resolve;
        }

    }
}