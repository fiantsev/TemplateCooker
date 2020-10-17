using Abstractions;
using System.Collections;
using System.Collections.Generic;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Utils;

namespace TemplateCooker.Service.Extraction
{
    public class MarkerExtractor : IMarkerExtractor, IEnumerable<Marker>
    {
        private readonly IEnumerable<ISheetAbstraction> _sheets;
        private readonly MarkerOptions _markerOptions;

        public MarkerExtractor(IWorkbookAbstraction workbook, MarkerOptions markerOptions)
        {
            _sheets = workbook.GetSheets();
            _markerOptions = markerOptions;
        }

        public MarkerExtractor(IEnumerable<ISheetAbstraction> sheets, MarkerOptions markerOptions)
        {
            _sheets = sheets;
            _markerOptions = markerOptions;
        }

        public MarkerExtractor(ISheetAbstraction sheet, MarkerOptions markerOptions)
        {
            _sheets = new List<ISheetAbstraction> { sheet };
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
                foreach (var row in sheet.GetUsedRows())
                {
                    foreach (var cell in row.GetUsedCells())
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
                                    SheetIndex = sheet.SheetIndex,
                                    RowIndex = row.FirstCell().RowIndex,
                                    ColumnIndex = cell.ColumnIndex
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