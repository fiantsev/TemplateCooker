using System;
using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.LayoutShifts;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker.Service.InjectionProcessing
{
    public class LayoutShiftProcessor : IInjectionProcessor
    {
        private List<LayoutShift> _layoutShifts;

        public LayoutShiftProcessor(IEnumerable<LayoutShift> layoutShifts)
        {
            _layoutShifts = layoutShifts.ToList();
        }


        public IEnumerable<InjectionContext> Process(IEnumerable<InjectionContext> originalInjectionContextStream)
        {
            ProcessRowShifts(originalInjectionContextStream);
            return originalInjectionContextStream; //всегда возвращаем оригинальный стрим
        }

        private List<RowLayoutShift> RecalculateOriginPositionsOnOneSheet(List<RowLayoutShift> rowLayoutShifts)
        {
            if (rowLayoutShifts.Select(x => x.OriginPosition.SheetIndex).Distinct().Count() != 1)
                throw new Exception("Все смещения должны быть на одном листе");

            var orderedRowLayoutShifts = rowLayoutShifts.OrderBy(x => x.OriginPosition.RowIndex).ToList();

            var updatedRowLayoutShifts = orderedRowLayoutShifts
                .Select(rowLayoutShift =>
                {
                    var accumulatedShiftFromPreviousShifts = orderedRowLayoutShifts.TakeWhile(x => x != rowLayoutShift).Sum(x => x.Amount);

                    return new RowLayoutShift
                    {
                        Amount = rowLayoutShift.Amount,
                        OriginPosition = new MarkerPosition
                        {
                            SheetIndex = rowLayoutShift.OriginPosition.SheetIndex,
                            RowIndex = rowLayoutShift.OriginPosition.RowIndex + accumulatedShiftFromPreviousShifts,
                            ColumnIndex = rowLayoutShift.OriginPosition.ColumnIndex,
                        }
                    };
                })
                .ToList();

            return updatedRowLayoutShifts;
        }

        private void ProcessRowShifts(IEnumerable<InjectionContext> injectionContextStream)
        {
            var rowLayoutShifts = _layoutShifts.OfType<RowLayoutShift>();
            var rowShiftsGroupedBySheets = rowLayoutShifts.GroupBy(x => x.OriginPosition.SheetIndex).ToList();

            rowShiftsGroupedBySheets.ForEach(rowShiftGroup =>
            {
                var sheetIndex = rowShiftGroup.Key;
                var recalculatedRowShifts = RecalculateOriginPositionsOnOneSheet(rowShiftGroup.ToList());

                var injectionsOnThisSheet = injectionContextStream
                    .Where(x => x.MarkerRange.StartMarker.Position.SheetIndex == sheetIndex)
                    //сортируем по индексу чтобы можно было отсекать по OriginPosition целые пачки инъекций
                    .OrderBy(x => x.MarkerRange.StartMarker.Position.RowIndex)
                    .ToList();

                recalculatedRowShifts
                    .ToList()
                    .ForEach(x =>
                        injectionsOnThisSheet
                            .SkipWhile(inj => inj.MarkerRange.StartMarker.Position.RowIndex <= x.OriginPosition.RowIndex)
                            .ToList()
                            .ForEach(inj => new MarkerRangeShifter().Shift(inj.MarkerRange, x))
                    );
            });
        }
    }
}
