namespace TemplateCooker.Service.InjectionProcessing
{
    //public class TableLayoutShiftProcessor : IInjectionProcessor
    //{
    //    //используется только здесь
    //    class MarkerGroupOnOneRow
    //    {
    //        public TableInjection TableInjection { get; set; }
    //        public SrcPosition StartMarkerPosition { get; set; }
    //        public string InjectionGuid { get; set; }
    //    }

    //    public IEnumerable<InjectionContext> Process(IEnumerable<InjectionContext> injectionContextStream)
    //    {
    //        var result = SynchronizeLayoutShifts(injectionContextStream, out var layoutShifts);

    //        var layoutShiftProcessor = new LayoutShiftProcessor(layoutShifts);

    //        var result2 = layoutShiftProcessor.Process(result);

    //        return result2;
    //    }

    //    /// <summary>
    //    /// Находит маркеры находящиеся на одной строке и требующие смещение контента
    //    /// Для всех таких маркеров, кроме первого, убираем необходимость смещать контент.
    //    /// И при этом первому выставляем количетсво элементов к смещению = максимальному количество данных пришедших в данной группе маркеров
    //    /// </summary>
    //    private IEnumerable<InjectionContext> SynchronizeLayoutShifts(
    //        IEnumerable<InjectionContext> injectionContextStream,
    //        out List<LayoutShift> layoutShifts
    //    )
    //    {
    //        var localLayoutShifts = new List<LayoutShift>();
    //        var cachedInjectionContexts = injectionContextStream.ToList();//кэшируем важно!

    //        //находим все инъекции табличного типа и активированным режимом смещения строк
    //        var tableInjections = cachedInjectionContexts
    //            .Select(context =>
    //            {
    //                if (context.Injection is TableInjection tableInjection)
    //                    return new MarkerGroupOnOneRow
    //                    {
    //                        TableInjection = tableInjection,
    //                        StartMarkerPosition = context.MarkerRange.StartMarker.Position,
    //                        InjectionGuid = context.Guid,
    //                    };

    //                return null;
    //            })
    //            .Where(x => x != null && x.TableInjection.LayoutShift == LayoutShiftType.MoveRows);

    //        //группируем инъекции по строкам
    //        var tableInjectionsGroupedByRow = tableInjections
    //            .GroupBy(x => new SrcPosition
    //            {
    //                SheetIndex = x.StartMarkerPosition.SheetIndex,
    //                RowIndex = x.StartMarkerPosition.RowIndex,
    //                ColumnIndex = 0
    //            });

    //        //группы содержащие один маркер
    //        var oneMarkerGroups = tableInjectionsGroupedByRow
    //            .Where(group => group.Count() == 1)
    //            .ToList();

    //        oneMarkerGroups.ForEach(tableInjectionOnRow =>
    //            {
    //                HandleRestOfItemsWithLayoutShift(injectionContextStream, tableInjectionOnRow);
    //                var singleInjectionInTheGroup = tableInjectionOnRow.Single();

    //                var shiftAmount = singleInjectionInTheGroup.TableInjection.Resource.Object.Count > 0
    //                ? singleInjectionInTheGroup.TableInjection.Resource.Object.Count - 1
    //                : 0;

    //                var firstInjectionInTheRow = injectionContextStream.First(cached => true
    //                    && cached.MarkerRange.StartMarker.Position.SheetIndex == singleInjectionInTheGroup.StartMarkerPosition.SheetIndex
    //                    && cached.MarkerRange.StartMarker.Position.RowIndex == singleInjectionInTheGroup.StartMarkerPosition.RowIndex
    //                );

    //                ((TableInjection)firstInjectionInTheRow.Injection).LayoutShift = LayoutShiftType.MoveRows;
    //                ((TableInjection)firstInjectionInTheRow.Injection).СountOfRowsToInsert = shiftAmount;
    //                localLayoutShifts.Add(new RowLayoutShift
    //                {
    //                    OriginPosition = firstInjectionInTheRow.MarkerRange.StartMarker.Clone().Position,
    //                    Amount = shiftAmount
    //                });
    //            });

    //        //группы содержащие несколько маркеров на одной строке
    //        var severalMarkerGroups = tableInjectionsGroupedByRow
    //            .Where(group => group.Count() > 1)
    //            .ToList();

    //        severalMarkerGroups.ForEach(group =>
    //            {
    //                HandleRestOfItemsWithLayoutShift(injectionContextStream, group);
    //                var layoutShiftForGroup = HandleFirstItemWithLayoutShift(injectionContextStream, group);
    //                localLayoutShifts.Add(layoutShiftForGroup);
    //            });

    //        layoutShifts = localLayoutShifts;
    //        return cachedInjectionContexts;
    //    }

    //    //для первого маркера в группе добавляем пустые строки (по максимальному количеству строк в группе)
    //    private LayoutShift HandleFirstItemWithLayoutShift(
    //        IEnumerable<InjectionContext> originalInjectionContextStream,
    //        IGrouping<SrcPosition, MarkerGroupOnOneRow> markerGroupOnOneRow
    //    )
    //    {
    //        var sheetIndex = markerGroupOnOneRow.Key.SheetIndex;
    //        var rowIndex = markerGroupOnOneRow.Key.RowIndex;

    //        var groupMaxDataRowCount = markerGroupOnOneRow.Max(x => x.TableInjection.Resource.Object.Count);
    //        var shiftAmount = groupMaxDataRowCount > 1
    //            ? groupMaxDataRowCount - 1
    //            : 0;

    //        var firstInjectionInTheRow = originalInjectionContextStream.First(inj => true
    //            && inj.MarkerRange.StartMarker.Position.SheetIndex == sheetIndex
    //            && inj.MarkerRange.StartMarker.Position.RowIndex == rowIndex
    //        );

    //        var tableInjectionWithRowsShifted = ((TableInjection)firstInjectionInTheRow.Injection);
    //        tableInjectionWithRowsShifted.LayoutShift = LayoutShiftType.MoveRows;
    //        tableInjectionWithRowsShifted.СountOfRowsToInsert = shiftAmount;

    //        return new RowLayoutShift
    //        {
    //            Amount = shiftAmount,
    //            OriginPosition = firstInjectionInTheRow.MarkerRange.StartMarker.Clone().Position
    //        };
    //    }

    //    //для последующих маркеров отключаем смещение, т.к. смещение будет выполнено один раз в первом маркере
    //    private void HandleRestOfItemsWithLayoutShift(
    //        IEnumerable<InjectionContext> originalInjectionContextStream,
    //        IEnumerable<MarkerGroupOnOneRow> markerGroupOnOneRow
    //    )
    //    {
    //        markerGroupOnOneRow
    //            .Select(x => originalInjectionContextStream.First(cached => cached.Guid == x.InjectionGuid))
    //            .ToList()
    //            .ForEach(x => (x.Injection as TableInjection).LayoutShift = LayoutShiftType.None);
    //    }
    //}
}