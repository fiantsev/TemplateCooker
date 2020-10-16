using System;
using System.Collections.Generic;

namespace TemplateCooker.Domain.ResourceObjects
{
    public class TableResourceObject : ResourceObject
    {
        public List<List<object>> Object { get; }

        public TableResourceObject(List<List<object>> table)
        {
            if (table == null)
                throw new NullReferenceException();

            Object = table;
        }
    }
}