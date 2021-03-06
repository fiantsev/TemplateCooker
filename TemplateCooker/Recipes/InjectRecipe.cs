﻿using PluginAbstraction;
using System;
using System.Collections.Generic;
using System.Linq;
using TemplateCooking.Domain.Injections;
using TemplateCooking.Domain.Markers;
using TemplateCooking.Service.Extraction;
using TemplateCooking.Service.InjectionProviders;
using TemplateCooking.Service.OperationExecutors;
using TemplateCooking.Service.Processing;

namespace TemplateCooking.Recipes
{
    public class InjectRecipe
    {
        public class Options
        {
            public IInjectionProcessor InjectionProcessor { get; set; }
            public IInjectionProvider InjectionProvider { get; set; }
            public MarkerOptions MarkerOptions { get; set; }

            public static IInjectionProcessor DefaultInjectionProcessor = new DefaultInjectionProcessor();
        }

        private Options _options;

        public InjectRecipe(Options options)
        {
            _options = options;
        }

        public void Cook(IWorkbookAbstraction workbook)
        {
            var injectionContexts = GenerateInjections(workbook);
            var processedInjectionContexts = ProcessInjections(workbook, injectionContexts);
            ExecuteInjections(workbook, processedInjectionContexts);
        }

        private List<InjectionContext> GenerateInjections(IWorkbookAbstraction workbook)
        {
            var markers = new MarkerExtractor(workbook, _options.MarkerOptions).GetMarkers();

            var markerRanges = new MarkerRangeCollection(markers);

            var injections = markerRanges
                .Select(markerRange => new InjectionContext
                {
                    MarkerRange = markerRange,
                    Injection = _options.InjectionProvider.Resolve(markerRange.StartMarker.Id),
                });

            return injections.ToList();
        }

        private List<AbstractOperation> ProcessInjections(IWorkbookAbstraction workbook, List<InjectionContext> injectionStream)
        {
            var processingStreams = _options.InjectionProcessor.Process(
                workbook,
                new ProcessingStreams { InjectionStream = injectionStream, OperationStream = new List<AbstractOperation>() }
            );
            return processingStreams.OperationStream;
        }

        private void ExecuteInjections(IWorkbookAbstraction workbook, List<AbstractOperation> operations)
        {
            foreach (var operation in operations)
            {
                var executor = GetExecutor(operation);
                executor.Execute(workbook, operation);
            }
        }

        private IOperationExecutor GetExecutor(AbstractOperation operation)
        {
            switch (operation)
            {
                case FillDownFormulas.Operation _: return new FillDownFormulas();
                case CopyPasteRowRangeWithStylesAndFormulas.Operation _: return new CopyPasteRowRangeWithStylesAndFormulas();
                case InsertEmptyRows.Operation _: return new InsertEmptyRows();
                case InsertImage.Operation _: return new InsertImage();
                case InsertTable.Operation _: return new InsertTable();
                case InsertText.Operation _: return new InsertText();
                default: throw new Exception($"Не найден исполнитель для операции {operation?.GetType().Name}");
            }
        }
    }
}