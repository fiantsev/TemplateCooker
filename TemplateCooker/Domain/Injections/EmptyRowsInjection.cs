namespace TemplateCooker.Domain.Injections
{
    public class EmptyRowsInjection : Injection
    {
        public EmptyRowsInjection(int rowsCount)
        {
            RowsCount = rowsCount;
        }

        public int RowsCount { get; set; }
    }
}