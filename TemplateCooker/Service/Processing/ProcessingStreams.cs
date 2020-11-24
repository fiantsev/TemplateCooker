using System.Collections.Generic;
using TemplateCooking.Service.OperationExecutors;
using TemplateCooking.Service.ResourceInjection;

namespace TemplateCooking.Service.Processing
{
    public class ProcessingStreams
    {
        public List<InjectionContext> InjectionStream { get; set; }
        public List<AbstractOperation> OperationStream { get; set; }
    }
}