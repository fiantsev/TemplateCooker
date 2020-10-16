using System.Collections.Generic;
using TemplateCooker.Domain.Markers;

namespace TemplateCooker.Service.Extraction
{
    public interface IMarkerExtractor
    {
        IEnumerable<Marker> GetMarkers();
    }
}