using TemplateCooker.Domain.Layout;

namespace TemplateCooker.Domain.Markers
{
    public class Marker
    {
        public string Id { get; set; }
        public SrcPosition Position { get; set; }
        public MarkerType MarkerType { get; set; }

        public Marker Clone()
        {
            return new Marker
            {
                Id = Id,
                MarkerType = MarkerType,
                Position = new SrcPosition(Position.SheetIndex, Position.RowIndex, Position.ColumnIndex)
            };
        }

    }
}