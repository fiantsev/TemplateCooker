using System.Collections.Generic;
using System.Linq;

namespace TemplateCooking.Domain.ResourceObjects
{
    public class TableWithHeaders
    {
        public List<List<object>> ColumnHeaders { get; set; }
        public List<List<object>> RowHeaders { get; set; }
        public List<List<object>> Body { get; set; }

        public TableWithHeaders()
        {
            ColumnHeaders = new List<List<object>>();
            RowHeaders = new List<List<object>>();
            Body = new List<List<object>>();
        }

        public TableWithHeaders(
            List<List<object>> columnHeaders,
            List<List<object>> rowHeaders,
            List<List<object>> body
        )
        {
            ColumnHeaders = columnHeaders ?? new List<List<object>>();
            RowHeaders = rowHeaders ?? new List<List<object>>();
            Body = body ?? new List<List<object>>();
        }

        public int Height => ColumnHeaders.Count + Body.Count;
        public int Width => (RowHeaders.FirstOrDefault()?.Count ?? 0) + (Body.FirstOrDefault()?.Count ?? 0);
    }
}