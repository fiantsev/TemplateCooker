using System;
using System.Linq;
using TemplateCooker.Domain.Injections;

namespace TemplateCooker.Service.ResourceInjection.Injectors
{
    public class ExtendFormulasDownInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => context =>
        {
            var injection = (context.Injection as ExtendFormulasDownInjection);
            var sheet = context.Workbook.GetSheet(injection.SheetIndex);

            var rowToCheckFormulas = sheet.GetRow(injection.FromRowIndex);

            var cellsWithFormula = rowToCheckFormulas.GetUsedCells().Where(x => x.HasFormula).ToList();

            for (var rowIndex = injection.FromRowIndex + 1; rowIndex <= injection.ToRowIndex; ++rowIndex)
            {
                var row = sheet.GetRow(rowIndex);
                foreach (var cellWithFormula in cellsWithFormula)
                    cellWithFormula.Copy(row.GetCell(cellWithFormula.ColumnIndex));
            }
        };
    }
}