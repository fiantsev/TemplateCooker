using System;

namespace TemplateCooker.Domain.Markers
{
    public class MarkerRange
    {
        public Marker StartMarker { get; }

        private Marker _EndMarker;
        public Marker EndMarker
        {
            get
            {
                throw new Exception("Using EndMarker is unsupported and therefore denied for now. It will be activated in the upcoming versions");
            }

            private set
            {
                _EndMarker = value;
            }
        }
        public bool Collapsed { get; }

        public MarkerRange(Marker startMarker, Marker endMarker = null)
        {
            if (endMarker == null)
            {
                endMarker = startMarker.Clone();
                endMarker.MarkerType = MarkerType.End;
            }

            if (startMarker.MarkerType != MarkerType.Start || endMarker.MarkerType != MarkerType.End || startMarker.Id != endMarker.Id)
                throw new ArgumentException();

            StartMarker = startMarker;
            EndMarker = endMarker;

            Collapsed = startMarker.Position == endMarker.Position;
        }
    }
}