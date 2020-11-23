using System;
using TemplateCooking.Domain.Layout;

namespace TemplateCooking.Domain.Markers
{
    public class Marker
    {
        public Marker(string id, SrcPosition position, MarkerType markerType)
        {
            Id = id ?? throw new ArgumentNullException(nameof(id));
            Position = position ?? throw new ArgumentNullException(nameof(position));
            MarkerType = markerType;
        }

        public string Id { get; }
        public SrcPosition Position { get; }
        public MarkerType MarkerType { get; }

        public Marker Clone()
        {
            return new Marker(Id, Position, MarkerType);
        }

        public Marker WithShift(int rowShift = 0, int columnShift = 0)
        {
            return new Marker(Id, Position.WithShift(rowShift, columnShift), MarkerType);
        }
    }
}