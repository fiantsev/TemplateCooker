using PluginAbstraction;
using System;
using System.Collections.Generic;
using System.Linq;
using TemplateCooking.Domain.Injections;
using TemplateCooking.Domain.Layout;
using TemplateCooking.Domain.Markers;
using TemplateCooking.Domain.ResourceObjects;
using TemplateCooking.Service.OperationExecutors;
using TemplateCooking.Service.Utils;

namespace TemplateCooking.Service.Processing
{
    public class DefaultInjectionProcessor : IInjectionProcessor
    {
        public ProcessingStreams Process(IWorkbookAbstraction workbook, ProcessingStreams processingStreams)
        {
            var injectionStream = processingStreams.InjectionStream;
            var operationStream = processingStreams.OperationStream;

            foreach (var injectionsOnOneSheet in injectionStream.GroupBy(x => x.MarkerRange.StartMarker.Position.SheetIndex))
            {
                //при вставление новых строк последующие операции должны смещаться пропорционально
                var rowShiftAccumulated = 0;
                var sheetIndex = injectionsOnOneSheet.First().MarkerRange.StartMarker.Position.SheetIndex;

                foreach (var injectionsOnOneRow in injectionsOnOneSheet.GroupBy(x => x.MarkerRange.StartMarker.Position.RowIndex))
                {
                    var rowIndex = injectionsOnOneRow.First().MarkerRange.StartMarker.Position.RowIndex;
                    var srPositionWithShift = new SrPosition(sheetIndex, rowIndex).WithShift(rowShiftAccumulated);
                    var (maxCellHeight, maxRowCount) = FindMaximums(workbook, injectionsOnOneRow.ToList());

                    //количество новых вставок необходимых произвести
                    var insertionCount = Math.Max(0, maxRowCount - 1);
                    //сколько пустых строк нужно вставить что бы все данные пришедшие из таблиц уместились
                    var rowCountToInsert = insertionCount * maxCellHeight;

                    if (rowCountToInsert > 0)
                    {
                        operationStream.Add(new InsertEmptyRows.Operation
                        {
                            Position = srPositionWithShift.WithShift(maxCellHeight - 1),
                            RowsCount = rowCountToInsert
                        });
                        operationStream.Add(new CopyPasteRowRangeWithStylesAndFormulas.Operation
                        {
                            CopyFromRow = srPositionWithShift,
                            CopyToRow = srPositionWithShift.WithShift(maxCellHeight - 1),
                            PasteStartRow = srPositionWithShift.WithShift(maxCellHeight - 1).WithShift(1),
                            PasteCount = insertionCount,
                        });
                        operationStream.AddRange(injectionsOnOneRow.Select(x => ConvertToOperation(x.Injection, x.MarkerRange.WithShift(rowShiftAccumulated), maxCellHeight)));
                        rowShiftAccumulated += rowCountToInsert;
                    }
                    else
                    {
                        operationStream.AddRange(injectionsOnOneRow.Select(x => ConvertToOperation(x.Injection, x.MarkerRange.WithShift(rowShiftAccumulated), maxCellHeight)));
                    }
                }
            }

            return processingStreams;
        }

        /// <summary>
        /// maxCellHeight - максимальная высота ячейки с маркером (смердженная ячейка может включать несколько строк в себя)
        /// maxRowCount - максимальное количество строк пришедших в табличных данных
        /// </summary>
        private (int maxCellHeight, int maxRowCount) FindMaximums(IWorkbookAbstraction workbook, List<InjectionContext> contexts)
        {
            //находим все таблицы со смещением строк (расположенные на одной строке)
            //их может и не быть ни одной
            var tablesWithMoveRows = contexts
                .Where(x => x.Injection is TableInjection tableInjection && tableInjection.LayoutShift == LayoutShiftType.MoveRows);

            if (tablesWithMoveRows.Count() == 0)
                return (0, 0);

            //находим высоту самого высокого маркера
            var maxCellHeight = tablesWithMoveRows
                .Max(x => workbook.GetCell(x.MarkerRange.StartMarker.Position).GetMergedRange().Height);

            //находим максимальное количество строк пришедших в данных какой либо из таблиц
            var maxRowCount = tablesWithMoveRows
                .Max(x => (x.Injection as TableInjection).Resource.Object.Count);

            return (maxCellHeight, maxRowCount);
        }


        /// <summary> TODO: избавиться от этой функции => для этог слить воедино (Injection и Injection.Operation) => в одну абстракцию </summary>
        private AbstractOperation ConvertToOperation(Injection injection, MarkerRange markerRange, int fixedRowStep)
        {
            switch (injection)
            {
                case TableInjection tableInjection:
                    return new InsertTable.Operation
                    {
                        Position = markerRange.StartMarker.Position,
                        Table = tableInjection.Resource.Object,
                        FixedRowStep = tableInjection.LayoutShift == LayoutShiftType.MoveRows
                            ? fixedRowStep
                            : 0
                    };
                case ImageInjection imageInjection:
                    return new InsertImage.Operation
                    {
                        Position = markerRange.StartMarker.Position,
                        Image = imageInjection.Resource.Object,
                    };
                case TextInjection textInjection:
                    return new InsertText.Operation
                    {
                        Position = markerRange.StartMarker.Position,
                        Text = textInjection.Resource.Object,
                    };
                default:
                    throw new Exception($"Неизвестный тип инъекции: {injection?.GetType().Name}");
            }
        }
    }
}