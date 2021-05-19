using System;
using System.Collections.Generic;

namespace TemplateCooking.Domain.ResourceObjects
{
    public class TableResourceObject : ResourceObject
    {
        public TableWithHeaders Object { get; }

        public TableResourceObject(List<List<object>> tableBody)
        {
            if (tableBody == null)
                throw new NullReferenceException();

            Object = new TableWithHeaders(
                new List<List<object>>(),
                new List<List<object>>(),
                tableBody
            );
        }

        public TableResourceObject(TableWithHeaders table)
        {
            if (table == null)
                throw new NullReferenceException();

            Object = table;
        }
    }
}