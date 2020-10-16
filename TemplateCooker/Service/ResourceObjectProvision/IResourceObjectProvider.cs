using System;
using TemplateCooker.Domain.ResourceObjects;

namespace TemplateCooker.Service.ResourceObjectProvision
{
    public interface IResourceObjectProvider
    {
        Func<string, ResourceObject> Resolve { get; }
    }
}