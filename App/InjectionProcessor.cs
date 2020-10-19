using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.InjectionProcessing;
using TemplateCooker.Service.ResourceInjection;

namespace App
{
    public class InjectionProcessor : IInjectionProcessor
    {
        public IEnumerable<InjectionContext> Process(IEnumerable<InjectionContext> injectionContexts)
        {
            var result = SynchronizeLayoutShifts(injectionContexts);
            return result;
        }

        /// <summary>
        /// Находит маркеры находящиеся на одной строке и требующие смещение контента
        /// Для всех таких маркеров, кроме первого, убираем необходимость смещать контент.
        /// И при этом первому выставляем количетсво элементов к смещению = максимальному количество данных пришедших в данной группе маркеров
        /// </summary>
        /// <param name="injectionContexts"></param>
        private IEnumerable<InjectionContext> SynchronizeLayoutShifts(IEnumerable<InjectionContext> injectionContexts)
        {
            return CalculateLayoutChanges(injectionContexts);

        }

        private IEnumerable<InjectionContext> CalculateLayoutChanges(IEnumerable<InjectionContext> injectionContexts)
        {
            var cachedInjectionContexts = injectionContexts.ToList();//кэшируем важно!

            var s1 = cachedInjectionContexts
                .Select(context =>
                {
                    if (context.Injection is TableInjection tableInjection)
                        return new
                        {
                            TableInjection = tableInjection,
                            StartMarkerPosition = context.MarkerRange.StartMarker.Position,
                            InjectionGuid = context.Guid,
                        };

                    return null;
                })
                .Where(x => x != null && x.TableInjection.LayoutShift == LayoutShiftType.MoveRows);

            var s2 = s1
                .GroupBy(x => new MarkerPosition
                {
                    SheetIndex = x.StartMarkerPosition.SheetIndex,
                    RowIndex = x.StartMarkerPosition.RowIndex,
                    ColumnIndex = 0
                })
                .Where(group => group.Count() > 1);

            s2
                .ToList()
                .ForEach(group =>
                {
                    var groupMaxDataRowCount = group.Max(x => x.TableInjection.Resource.Object.Count);

                    var firstInjectionByOriginalOrder = cachedInjectionContexts.First(x => x.Guid == group.First().InjectionGuid);
                    var firstTableObject = ((TableInjection)firstInjectionByOriginalOrder.Injection).Resource.Object;
                    var countOfEmptyStringShouldBeAdded = groupMaxDataRowCount - firstTableObject.Count;

                    AddEmptyRows(firstTableObject, countOfEmptyStringShouldBeAdded);

                    group
                        .Skip(1)
                        .Select(x => cachedInjectionContexts.First(cached => cached.Guid == x.InjectionGuid))
                        .ToList()
                        .ForEach(x => (x.Injection as TableInjection).LayoutShift = LayoutShiftType.None);
                });

            return cachedInjectionContexts;
        }


        private void AddEmptyRows(List<List<object>> table, int count)
        {
            var addition = Enumerable.Repeat(Enumerable.Repeat(string.Empty, table[0].Count).Cast<object>().ToList(), count);
            table.AddRange(addition);
        }
    }
}