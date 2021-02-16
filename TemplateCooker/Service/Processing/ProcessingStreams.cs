using System.Collections.Generic;
using TemplateCooking.Domain.Injections;
using TemplateCooking.Service.OperationExecutors;

namespace TemplateCooking.Service.Processing
{
    public class ProcessingStreams
    {
        public List<InjectionContext> InjectionStream { get; set; }
        public List<AbstractOperation> OperationStream { get; set; }
    }
}