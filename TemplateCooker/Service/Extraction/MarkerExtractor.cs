using ClosedXML.Excel;
using System.Collections;
using System.Collections.Generic;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Utils;

namespace TemplateCooker.Service.Extraction
{
    public class MarkerExtractor : IMarkerExtractor, IEnumerable<Marker>
    {
        private readonly IEnumerable<IXLWorksheet> _sheets;
        private readonly MarkerOptions _markerOptions;

        public MarkerExtractor(IXLWorkbook workbook, MarkerOptions markerOptions)
        {
            _sheets = workbook.Worksheets;
            _markerOptions = markerOptions;
        }

        public MarkerExtractor(IEnumerable<IXLWorksheet> sheets, MarkerOptions markerOptions)
        {
            _sheets = sheets;
            _markerOptions = markerOptions;
        }

        public MarkerExtractor(IXLWorksheet sheet, MarkerOptions markerOptions)
        {
            _sheets = new List<IXLWorksheet> { sheet };
            _markerOptions = markerOptions;
        }

        public IEnumerable<Marker> GetMarkers()
        {
            return this;
        }

        IEnumerator<Marker> IEnumerable<Marker>.GetEnumerator()
        {
            foreach (var sheet in _sheets)
            {
                var rangeUsed = sheet.RangeUsed();
                if (rangeUsed == null) continue;

                foreach (var row in rangeUsed.Rows())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (CellUtils.IsMarkedCell(cell, _markerOptions))
                        {
                            var markerId = CellUtils.ExtractMarkerValue(cell, _markerOptions);
                            var isEndMarker = markerId.Substring(0, _markerOptions.Terminator.Length) == _markerOptions.Terminator;
                            var marker = new Marker
                            {
                                Id = isEndMarker
                                    ? markerId.Substring(_markerOptions.Terminator.Length)
                                    : markerId,
                                Position = new MarkerPosition
                                {
                                    SheetIndex = SheetUtils.SheetIndex(sheet),
                                    RowIndex = row.FirstCell().Address.RowNumber,
                                    CellIndex = cell.Address.ColumnNumber
                                },
                                MarkerType = isEndMarker ? MarkerType.End : MarkerType.Start
                            };
                            yield return marker;
                        }
                    }
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<Marker>)this).GetEnumerator();
        }
    }
}