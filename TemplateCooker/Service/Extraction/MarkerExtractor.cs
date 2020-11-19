using PluginAbstraction;
using System.Collections;
using System.Collections.Generic;
using TemplateCooker.Domain.Layout;
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
                        var markerId = CellUtils.ExtractMarkerValueOrNull(cell, _markerOptions);
                        if (markerId == null)
                            continue;
                        var isEndMarker = markerId.Substring(0, _markerOptions.Terminator.Length) == _markerOptions.Terminator;
                        var marker = new Marker(
                            id: isEndMarker
                                ? markerId.Substring(_markerOptions.Terminator.Length)
                                : markerId,
                            position: new SrcPosition(sheet.SheetIndex, row.FirstCell().RowIndex, cell.ColumnIndex),
                            markerType: isEndMarker ? MarkerType.End : MarkerType.Start
                        );
                        yield return marker;
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