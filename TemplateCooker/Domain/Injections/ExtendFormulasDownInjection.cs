namespace TemplateCooker.Domain.Injections
{
    public class ExtendFormulasDownInjection : Injection
    {
        public int SheetIndex { get; set; }
        public int FromRowIndex { get; set; }
        public int ToRowIndex { get; set; }
    }
}