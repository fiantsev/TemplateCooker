using System;
using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.Layout;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker.Service.Layout
{
    public class LayoutService
    {
        class LayoutShift
        {
            public LayoutShiftIntent Intent { get; set; }
            public InjectionContext InjectionInitiator { get; set; }
        }

        /// <summary>
        /// инъекции должны приходить по одному листу
        /// </summary>
        public List<InjectionContext> ProcessLayout(List<InjectionContext> contexts)
        {
            if (contexts.Select(x => x.MarkerRange.StartMarker.Position.SheetIndex).Distinct().Count() > 1)
                throw new Exception("инъекции должны приходить по одному листу");
            var s1 = ExtractLayoutIntents(contexts);
            var s2 = InnerProcessLayout(s1);
            return s2;
        }

        //one-to-one mapping should be preserved
        private List<LayoutShift> ExtractLayoutIntents(List<InjectionContext> contexts)
        {
            return contexts.ToDictionary(context => context, context =>
                {
                    var rowIndex = context.MarkerRange.StartMarker.Position.RowIndex;
                    var columnIndex = context.MarkerRange.StartMarker.Position.ColumnIndex;
                    var rcPosition = new RcPosition(rowIndex, columnIndex);

                    switch (context.Injection)
                    {
                        case TableInjection tableInjection:
                            {
                                var tableRowCount = tableInjection.Resource.Object.Count;
                                var tableColumnCount = tableRowCount == 0 ? 0 : tableInjection.Resource.Object[0].Count;
                                var rcDimensions = new RcDimensions(
                                    Math.Max(1, tableRowCount),
                                    Math.Max(1, tableColumnCount)
                                );
                                return new LayoutShiftIntent(new LayoutElement(rcPosition, rcDimensions), tableInjection.LayoutShift);
                            }
                        default:
                            return new LayoutShiftIntent(new LayoutElement(rcPosition, new RcDimensions(1, 1)), LayoutShiftType.None);
                    }
                })
                .Select(x => new LayoutShift { Intent = x.Value, InjectionInitiator = x.Key })
                .ToList();
        }

        /// <summary>
        /// метод должен вернуть обновленные позиции инъекций и дополнить поток инъекций новыми утилити-инъекциями (напр добавление пустых строк)
        /// </summary>
        /// <param name="layoutMappings"></param>
        /// <returns></returns>
        private List<InjectionContext> InnerProcessLayout(List<LayoutShift> layoutShifts)
        {
            var outputStream = new List<InjectionContext>();


            var moveRows = layoutShifts
                .Where(x => x.Intent.Type == LayoutShiftType.MoveRows);

            var moveCells = layoutShifts
                .Where(x => x.Intent.Type == LayoutShiftType.MoveCells);

            //layoutMappings
            //    .Select(x => new
            //    {
            //        //OutputStream = outputStream,
            //        InjectionContext = x.Key,
            //        LayoutShiftIntent = x.Value,
            //    })
            //    .ToList()
            //    .ForEach(x =>
            //    {

            //    });
            //var s1 = layoutMappings.ToList();
            //var s2 = s1.
        }

        private void ProcessMoveRows()
        {

        }

    }
}
