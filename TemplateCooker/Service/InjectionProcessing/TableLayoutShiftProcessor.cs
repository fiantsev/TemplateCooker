using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.LayoutShifts;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker.Service.InjectionProcessing
{
    public class TableLayoutShiftProcessor : IInjectionProcessor
    {
        //private LayoutShiftProcessor _layoutShiftProcessor = new LayoutShiftProcessor();
        //используется только здесь
        class MarkerGroupOnOneRow
        {
            public TableInjection TableInjection { get; set; }
            public MarkerPosition StartMarkerPosition { get; set; }
            public string InjectionGuid { get; set; }
        }

        public IEnumerable<InjectionContext> Process(IEnumerable<InjectionContext> injectionContextStream)
        {
            var result = SynchronizeLayoutShifts(injectionContextStream, out var layoutShifts);

            var layoutShiftProcessor = new LayoutShiftProcessor(layoutShifts);

            var result2 = layoutShiftProcessor.Process(result);

            return result2;
        }

        /// <summary>
        /// Находит маркеры находящиеся на одной строке и требующие смещение контента
        /// Для всех таких маркеров, кроме первого, убираем необходимость смещать контент.
        /// И при этом первому выставляем количетсво элементов к смещению = максимальному количество данных пришедших в данной группе маркеров
        /// </summary>
        /// <param name="injectionContexts"></param>
        private IEnumerable<InjectionContext> SynchronizeLayoutShifts(
            IEnumerable<InjectionContext> injectionContextStream,
            out List<LayoutShift> layoutShifts
        )
        {
            var localLayoutShifts = new List<LayoutShift>();
            var cachedInjectionContexts = injectionContextStream.ToList();//кэшируем важно!

            //находим все инъекции табличного типа и активированным режимом смещения строк
            var tableInjections = cachedInjectionContexts
                .Select(context =>
                {
                    if (context.Injection is TableInjection tableInjection)
                        return new MarkerGroupOnOneRow
                        {
                            TableInjection = tableInjection,
                            StartMarkerPosition = context.MarkerRange.StartMarker.Position,
                            InjectionGuid = context.Guid,
                        };

                    return null;
                })
                .Where(x => x != null && x.TableInjection.LayoutShift == LayoutShiftType.MoveRows);

            //группируем инъекции по строкам
            var tableInjectionsGroupedByRow = tableInjections
                .GroupBy(x => new MarkerPosition
                {
                    SheetIndex = x.StartMarkerPosition.SheetIndex,
                    RowIndex = x.StartMarkerPosition.RowIndex,
                    ColumnIndex = 0
                });

            //группы содержащие один маркер
            var oneMarkerGroups = tableInjectionsGroupedByRow
                .Where(group => group.Count() == 1)
                .ToList();

            oneMarkerGroups.ForEach(tableInjectionOnRow =>
                {
                    localLayoutShifts.Add(new RowLayoutShift
                    {
                        OriginPosition = tableInjectionOnRow.Single().StartMarkerPosition,
                        Amount = tableInjectionOnRow.Single().TableInjection.Resource.Object.Count - 1
                    });
                });

            //группы содержащие несколько маркеров на одной строке
            var severalMarkerGroups = tableInjectionsGroupedByRow
                .Where(group => group.Count() > 1)
                .ToList();

            severalMarkerGroups.ForEach(group =>
                {
                    var layoutShiftForGroup = HandleFirstItemWithLayoutShift(injectionContextStream, group);
                    localLayoutShifts.Add(layoutShiftForGroup);
                    HandleRestOfItemsWithLayoutShift(injectionContextStream, group);
                });

            layoutShifts = localLayoutShifts;
            return cachedInjectionContexts;
        }

        private void AddEmptyRows(List<List<object>> table, int count)
        {
            var addition = Enumerable.Repeat(Enumerable.Repeat(string.Empty, table[0].Count).Cast<object>().ToList(), count);
            table.AddRange(addition);
        }

        //для первого маркера в группе добавляем пустые строки (по максимальному количеству строк в группе)
        private LayoutShift HandleFirstItemWithLayoutShift(
            IEnumerable<InjectionContext> originalInjectionContextStream,
            IEnumerable<MarkerGroupOnOneRow> markerGroupOnOneRow
        )
        {
            var groupMaxDataRowCount = markerGroupOnOneRow.Max(x => x.TableInjection.Resource.Object.Count);
            var firstInjectionByOriginalOrder = originalInjectionContextStream.First(x => markerGroupOnOneRow.Any(g => g.InjectionGuid == x.Guid));
            var firstTableObject = ((TableInjection)firstInjectionByOriginalOrder.Injection).Resource.Object;
            var countOfEmptyStringShouldBeAdded = groupMaxDataRowCount - firstTableObject.Count;
            AddEmptyRows(firstTableObject, countOfEmptyStringShouldBeAdded);
            return new RowLayoutShift
            {
                Amount = groupMaxDataRowCount - 1,
                OriginPosition = firstInjectionByOriginalOrder.MarkerRange.StartMarker.Clone().Position
            };
        }

        //для последующих маркеров отключаем смещение, т.к. смещение будет выполнено один раз в первом маркере
        private void HandleRestOfItemsWithLayoutShift(
            IEnumerable<InjectionContext> originalInjectionContextStream,
            IEnumerable<MarkerGroupOnOneRow> markerGroupOnOneRow
        )
        {
            var firstInjectionByOriginalOrder = originalInjectionContextStream.First(x => markerGroupOnOneRow.Any(g => g.InjectionGuid == x.Guid));
            markerGroupOnOneRow
                .Where(x => x.InjectionGuid != firstInjectionByOriginalOrder.Guid)
                .Select(x => originalInjectionContextStream.First(cached => cached.Guid == x.InjectionGuid))
                .ToList()
                .ForEach(x => (x.Injection as TableInjection).LayoutShift = LayoutShiftType.None);

        }
    }
}