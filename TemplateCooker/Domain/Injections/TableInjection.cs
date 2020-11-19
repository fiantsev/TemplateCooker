using TemplateCooker.Domain.ResourceObjects;

namespace TemplateCooker.Domain.Injections
{
    public class TableInjection : Injection
    {
        public TableResourceObject Resource { get; set; }

        public LayoutShiftType LayoutShift { get; set; }

        /// <summary>
        /// HACK: переделать механизм
        /// УДАЛИТЬ !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        /// </summary>
        public int СountOfRowsToInsert { get; set; } = 0;
    }
}