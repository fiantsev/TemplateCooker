using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace TemplateCooker.Domain.Markers
{
    public class MarkerRangeCollection : IEnumerable<MarkerRange>
    {
        private readonly IEnumerable<Marker> _markers;

        public MarkerRangeCollection(IEnumerable<Marker> markers)
        {
            _markers = markers;
        }

        public IEnumerator<MarkerRange> GetEnumerator()
        {
            //кэшируем
            var endMarkers = _markers.Where(x => x.MarkerType == MarkerType.End).ToList();
            var startMarkers = _markers.Where(x => x.MarkerType == MarkerType.Start);

            foreach (var startMarker in startMarkers)
            {
                var endMarker = endMarkers.FirstOrDefault(x => x.Id == startMarker.Id);
                var markerRegion = new MarkerRange(startMarker, endMarker);
                yield return markerRegion;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}