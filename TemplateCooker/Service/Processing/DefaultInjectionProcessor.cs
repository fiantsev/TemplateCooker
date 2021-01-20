using System;
using System.Collections.Generic;
using System.Linq;
using TemplateCooking.Domain.Injections;
using TemplateCooking.Domain.Layout;
using TemplateCooking.Domain.Markers;
using TemplateCooking.Domain.ResourceObjects;
using TemplateCooking.Service.OperationExecutors;
using TemplateCooking.Service.ResourceInjection;
using TemplateCooking.Service.Utils;
using PluginAbstraction;

namespace TemplateCooking.Service.Processing
{
    public class DefaultInjectionProcessor : IInjectionProcessor
    {
        public ProcessingStreams Process(IWorkbookAbstraction workbook,  ProcessingStreams processingStreams)
        {
            if (processingStreams.InjectionStream.Select(x => x.MarkerRange.StartMarker.Position.SheetIndex).Distinct().Count() > 1)
                throw new Exception("инъекции должны приходить по одному листу");

            InnerProcessLayout(workbook, processingStreams);

            return processingStreams;
        }


        private void InnerProcessLayout(IWorkbookAbstraction workbook, ProcessingStreams processingStreams)
        {
            var injectionStream = processingStreams.InjectionStream;
            var operationStream = processingStreams.OperationStream;
            //при вставление новых строк последующие операции должны смещаться пропорционально
            var previousRowShiftAmountAccumulated = 0;

            injectionStream
                .GroupBy(x => x.MarkerRange.StartMarker.Position.RowIndex)
                .ToList()
                .ForEach(contextsOnSameRow =>
                {
                    var firstContext = contextsOnSameRow.First();
                    var sheetIndex = firstContext.MarkerRange.StartMarker.Position.SheetIndex;
                    var rowIndex = firstContext.MarkerRange.StartMarker.Position.RowIndex;

                    //HACK: refactoring
                    var (rowCountToInsert, rowToInsertPosition, startMarkerPosition, pasteCount) = FindHowMuchNewEmptyRowsToInsertAndWhere(workbook, contextsOnSameRow.ToList());

                    if (rowCountToInsert > 0)
                    {
                        operationStream.Add(new InsertEmptyRows.Operation { Position = rowToInsertPosition.ToSrPosition().WithShift(previousRowShiftAmountAccumulated), RowsCount = rowCountToInsert });
                        operationStream.Add(new CopyPasteRowRange.Operation {
                            CopyFromRow = new SrPosition(sheetIndex, startMarkerPosition.RowIndex + previousRowShiftAmountAccumulated),
                            CopyToRow = new SrPosition(sheetIndex, startMarkerPosition.RowIndex + workbook.GetCell(startMarkerPosition).GetMergedRange().Height - 1 + previousRowShiftAmountAccumulated),
                            PasteStartRow = rowToInsertPosition.ToSrPosition().WithShift(previousRowShiftAmountAccumulated+1),
                            PasteCount = pasteCount,
                        });
                        operationStream.AddRange(contextsOnSameRow.Select(x => ConvertToOperation(x.Injection, x.MarkerRange.WithShift(previousRowShiftAmountAccumulated))));
                        //operationStream.Add(new FillDownFormulas.Operation { From = new SrPosition(sheetIndex, rowIndex + previousRowShiftAmountAccumulated), To = new SrPosition(sheetIndex, rowIndex + rowCountToInsert + previousRowShiftAmountAccumulated) });
                        previousRowShiftAmountAccumulated += rowCountToInsert;
                    }
                    else
                    {
                        operationStream.AddRange(contextsOnSameRow.Select(x => ConvertToOperation(x.Injection, x.MarkerRange.WithShift(previousRowShiftAmountAccumulated))));
                    }
                });
        }

        //HACK: refactoring return type to seprate class
        private (int, SrcPosition, SrcPosition, int) FindHowMuchNewEmptyRowsToInsertAndWhere(IWorkbookAbstraction workbook, List<InjectionContext> contexts)
        {
            var tables = contexts
                .Where(x =>
                {
                    if (x.Injection is TableInjection tableInjection && tableInjection.LayoutShift == LayoutShiftType.MoveRows)
                        return true;
                    else
                        return false;
                })
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
            return (countOfNewRowsToInsert, positionToInsertNewRows, max.Context.MarkerRange.StartMarker.Position, max.RowCount - 1 );
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