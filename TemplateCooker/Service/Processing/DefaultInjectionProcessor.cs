using System;
using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.Layout;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.OperationExecutors;
using TemplateCooker.Service.ResourceInjection;
using TemplateCooker.Service.Utils;

namespace TemplateCooker.Service.Processing
{
    public class DefaultInjectionProcessor : IInjectionProcessor
    {
        public void Process(List<InjectionContext> injectionStream, List<AbstractOperation> operationStream)
        {
            if (injectionStream.Select(x => x.MarkerRange.StartMarker.Position.SheetIndex).Distinct().Count() > 1)
                throw new Exception("инъекции должны приходить по одному листу");

            InnerProcessLayout(injectionStream, operationStream);
        }


        private void InnerProcessLayout(List<InjectionContext> injectionStream, List<AbstractOperation> operationStream)
        {
            //при вставление новых строк последующие операции должны смещаться пропорционально
            var previousRowShiftAmountAccumulated = 0;

            injectionStream
                .GroupBy(x => x.MarkerRange.StartMarker.Position.RowIndex)
                .ToList()
                .ForEach(contextsOnSameRow =>
                {
                    var firstContext = contextsOnSameRow.First();
                    var workbook = firstContext.Workbook;
                    var sheetIndex = firstContext.MarkerRange.StartMarker.Position.SheetIndex;
                    var rowIndex = firstContext.MarkerRange.StartMarker.Position.RowIndex;

                    var (rowCountToInsert, rowToInsertPosition) = FindHowMuchNewEmptyRowsToInsertAndWhere(contextsOnSameRow.ToList());

                    if (rowCountToInsert > 0)
                    {
                        operationStream.Add(new InsertEmptyRows.Operation { Position = rowToInsertPosition.ToSrPosition().WithShift(previousRowShiftAmountAccumulated), RowsCount = rowCountToInsert });
                        operationStream.AddRange(contextsOnSameRow.Select(x => ConvertToOperation(x.Injection, x.MarkerRange.WithShift(previousRowShiftAmountAccumulated))));
                        operationStream.Add(new FillDownFormulas.Operation { From = new SrPosition(sheetIndex, rowIndex + previousRowShiftAmountAccumulated), To = new SrPosition(sheetIndex, rowIndex + rowCountToInsert + previousRowShiftAmountAccumulated) });

                        previousRowShiftAmountAccumulated += rowCountToInsert;
                    }
                    else
                    {
                        operationStream.AddRange(contextsOnSameRow.Select(x => ConvertToOperation(x.Injection, x.MarkerRange.WithShift(previousRowShiftAmountAccumulated))));
                    }
                });
        }

        private (int, SrcPosition) FindHowMuchNewEmptyRowsToInsertAndWhere(List<InjectionContext> contexts)
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
                    StartMarkerCellHeight = x.Workbook.GetCell(x.MarkerRange.StartMarker.Position).GetMergedRange().Height,
                    RowCount = (x.Injection as TableInjection).Resource.Object.Count
                });

            var max = tables
                .OrderByDescending(x => x.StartMarkerCellHeight * x.RowCount)
                .FirstOrDefault();

            if (max == null)
                return (0, null);

            var countOfNewRowsToInsert = Math.Max(0, max.RowCount - 1) * max.StartMarkerCellHeight;
            var positionToInsertNewRows = max.Context.MarkerRange.StartMarker.Position.WithShift(max.StartMarkerCellHeight - 1, 0);
            return (countOfNewRowsToInsert, positionToInsertNewRows);
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