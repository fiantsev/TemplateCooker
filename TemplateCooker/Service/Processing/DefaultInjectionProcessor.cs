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
                    var (rowCountToInsert, rowToInsertPosition, startMarkerPosition, pasteCount) = FindHowMuchNewEmptyRowsToInsertAndWhere(workbook, injectionsOnOneRow.ToList());

                    if (rowCountToInsert > 0)
                    {
                        var copyFromRow = startMarkerPosition.ToSrPosition().WithShift(rowShiftAccumulated);
                        var copyToRow = copyFromRow.WithShift(workbook.GetCell(startMarkerPosition).GetMergedRange().Height - 1);
                        operationStream.Add(new InsertEmptyRows.Operation
                        {
                            Position = rowToInsertPosition.ToSrPosition().WithShift(rowShiftAccumulated),
                            RowsCount = rowCountToInsert
                        });
                        operationStream.Add(new CopyPasteRowRange.Operation
                        {
                            CopyFromRow = copyFromRow,
                            CopyToRow = copyToRow,
                            PasteStartRow = rowToInsertPosition.ToSrPosition().WithShift(rowShiftAccumulated + 1),
                            PasteCount = pasteCount,
                        });
                        operationStream.AddRange(injectionsOnOneRow.Select(x => ConvertToOperation(x.Injection, x.MarkerRange.WithShift(rowShiftAccumulated))));
                        rowShiftAccumulated += rowCountToInsert;
                    }
                    else
                    {
                        operationStream.AddRange(injectionsOnOneRow.Select(x => ConvertToOperation(x.Injection, x.MarkerRange.WithShift(rowShiftAccumulated))));
                    }
                }
            }

            return processingStreams;
        }

        /// <summary>
        /// rowCountToInsert - количество строк которое необходимо вставить
        /// rowToInsertPosition - позиция строки после которой необходимо вставить пустые строки
        /// startMarkerPosition - позиция маркера для которого необходимо стили строк продублировать ниже
        /// pasteCount - количество копипаста строк для сохранения стиля
        /// </summary>
        private (int rowCountToInsert, SrcPosition rowToInsertPosition, SrcPosition startMarkerPosition, int pasteCount)
            FindHowMuchNewEmptyRowsToInsertAndWhere(IWorkbookAbstraction workbook, List<InjectionContext> contexts)
        {
            var tables = contexts
                .Where(x => x.Injection is TableInjection tableInjection && tableInjection.LayoutShift == LayoutShiftType.MoveRows)
                .Select(x => new
                {
                    Context = x,
                    StartMarkerCellHeight = workbook.GetCell(x.MarkerRange.StartMarker.Position).GetMergedRange().Height,
                    RowCount = (x.Injection as TableInjection).Resource.Object.Count
                });

            var max = tables
                .OrderByDescending(x => x.StartMarkerCellHeight * x.RowCount)
                .FirstOrDefault();

            if (max == null)
                return (0, null, null, 0);

            var countOfNewRowsToInsert = Math.Max(0, max.RowCount - 1) * max.StartMarkerCellHeight;
            var positionToInsertNewRows = max.Context.MarkerRange.StartMarker.Position.WithShift(max.StartMarkerCellHeight - 1, 0);
            return (countOfNewRowsToInsert, positionToInsertNewRows, max.Context.MarkerRange.StartMarker.Position, max.RowCount - 1);
        }

        private AbstractOperation ConvertToOperation(Injection injection, MarkerRange markerRange)
        {
            switch (injection)
            {
                case TableInjection tableInjection:
                    return new InsertTable.Operation
                    {
                        Position = markerRange.StartMarker.Position,
                        Table = tableInjection.Resource.Object,
                        PreserveStyleOfFirstCell = tableInjection.LayoutShift == LayoutShiftType.MoveRows
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