namespace TemplateCooker.Domain.Markers
{
    public class Marker
    {
        public string Id { get; set; }
        public MarkerPosition Position { get; set; }
        public MarkerType MarkerType { get; set; }

        public Marker Clone()
        {
            return new Marker
            {
                Id = Id,
                MarkerType = MarkerType,
                Position = new MarkerPosition
                {
                    SheetIndex = Position.SheetIndex,
                    RowIndex = Position.RowIndex,
                    CellIndex = Position.CellIndex,
                }
            };
        }

    }
}