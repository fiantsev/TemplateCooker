using System.Collections.Generic;
using TemplateCooker.Service.Operations;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker.Service.Processing
{
    public interface IInjectionProcessor
    {
        /// <summary>
        /// перерабатывает входной поток инъекций в выходной поток операций
        /// например одна инъекция на вставление таблицы может превратиться в несколько операций (непосредственно вставление таблицы и вставление пустых строк при динамическом количестве строк)
        /// </summary>
        void Process(List<InjectionContext> injectionStream, List<AbstractOperation> operationStream);
    }
}