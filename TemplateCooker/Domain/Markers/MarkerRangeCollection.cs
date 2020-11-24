using System.Collections;
using System.Collections.Generic;

namespace TemplateCooking.Domain.Markers
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
            foreach (var marker in _markers)
                yield return new MarkerRange(marker);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}