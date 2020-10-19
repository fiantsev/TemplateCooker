namespace PluginAbstraction
{
    public class CustomProperties
    {
        public WorkbookProperties WorkbookProperties { get; set; }
        public bool RecalculateFormulasOnSave { get; set; }
    }

    public class WorkbookProperties
    {
        public bool ForceFullCalculation { get; set; }
        public bool FullCalculationOnLoad { get; set; }
    }
}