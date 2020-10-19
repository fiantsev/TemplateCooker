﻿using PluginAbstraction;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.Markers;

namespace TemplateCooker.Service.ResourceInjection
{
    public class InjectionContext
    {
        public InjectionContext()
        {
            Guid = System.Guid.NewGuid().ToString();
        }
        public string Guid { get; }
        public MarkerRange MarkerRange { get; set; }
        public Injection Injection { get; set; }
        public IWorkbookAbstraction Workbook { get; set; }
    }
}