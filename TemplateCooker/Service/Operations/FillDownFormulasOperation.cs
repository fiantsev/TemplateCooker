using PluginAbstraction;
using System.Linq;
using TemplateCooker.Domain.Layout;

namespace TemplateCooker.Service.Operations
{
    public class FillDownFormulasOperation : IOperation
    {
        public class Options
        {
            public SrPosition From { get; set; }
            public SrPosition To { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Options)untypedOptions;
            var from = options.From;
            var to = options.To;
            var sheet = workbook.GetSheet(from.SheetIndex);

            var rowToCheckFormulas = sheet.GetRow(from.RowIndex);

            var cellsWithFormula = rowToCheckFormulas.GetUsedCells().Where(x => x.HasFormula).ToList();

            for (var rowIndex = from.RowIndex + 1; rowIndex <= to.RowIndex; ++rowIndex)
            {
                var row = sheet.GetRow(rowIndex);
                foreach (var cellWithFormula in cellsWithFormula)
                    cellWithFormula.GetMergedRange().CopyTo(row.GetCell(cellWithFormula.ColumnIndex));
            }
        }
    }
}