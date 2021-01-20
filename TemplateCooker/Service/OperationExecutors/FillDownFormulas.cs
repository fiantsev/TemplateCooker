using PluginAbstraction;
using System.Linq;
using TemplateCooking.Domain.Layout;

namespace TemplateCooking.Service.OperationExecutors
{
    public class FillDownFormulas : IOperationExecutor
    {
        public class Operation : AbstractOperation
        {
            public SrPosition From { get; set; }
            public SrPosition To { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Operation)untypedOptions;
            var from = options.From;
            var to = options.To;
            var sheet = workbook.GetSheet(from.SheetIndex);

            var rowToCheckFormulas = sheet.GetRow(from.RowIndex);

            var cellsWithFormula = rowToCheckFormulas.GetUsedCells(false).Where(x => x.HasFormula).ToList();

            for (var rowIndex = from.RowIndex + 1; rowIndex <= to.RowIndex; ++rowIndex)
            {
                var row = sheet.GetRow(rowIndex);
                foreach (var cellWithFormula in cellsWithFormula)
                    cellWithFormula.GetMergedRange().CopyTo(row.GetCell(cellWithFormula.ColumnIndex));
            }
        }
    }
}