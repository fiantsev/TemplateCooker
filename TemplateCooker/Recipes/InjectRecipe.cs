using PluginAbstraction;
using System;
using System.Collections.Generic;
using System.Linq;
using TemplateCooking.Service.Processing;
using TemplateCooking.Domain.Markers;
using TemplateCooking.Service.Extraction;
using TemplateCooking.Service.InjectionProviders;
using TemplateCooking.Service.OperationExecutors;
using TemplateCooking.Service.Processing;
using TemplateCooking.Service.ResourceInjection;

namespace TemplateCooking.Recipes
{
    public class InjectRecipe
    {
        public class Options
        {
            public IWorkbookAbstraction Workbook { get; set; }
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

        public void Cook()
        {
            InjectDataSheetBySheet();
        }

        public void InjectDataSheetBySheet()
        {
            IWorkbookAbstraction workbook = _options.Workbook;

            foreach (var sheet in workbook.GetSheets())
            {
                var injectionContexts = GenerateInjections(workbook, sheet);
                var processedInjectionContexts = ProcessInjections(injectionContexts);
                ExecuteInjections(workbook, processedInjectionContexts);
            }
        }

        private List<InjectionContext> GenerateInjections(IWorkbookAbstraction workbook, ISheetAbstraction sheet)
        {
            var markers = new MarkerExtractor(sheet, _options.MarkerOptions).GetMarkers();

            var markerRanges = new MarkerRangeCollection(markers);

            var injections = markerRanges
                .Select(markerRange => new InjectionContext
                {
                    MarkerRange = markerRange,
                    Injection = _options.InjectionProvider.Resolve(markerRange.StartMarker.Id),
                    Workbook = workbook
                });

            return injections.ToList();
        }

        private List<AbstractOperation> ProcessInjections(List<InjectionContext> injectionStream)
        {
            var processingStreams = _options.InjectionProcessor.Process(new ProcessingStreams
            {
                InjectionStream = injectionStream,
                OperationStream = new List<AbstractOperation>(),
            });
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
                case InsertEmptyRows.Operation _: return new InsertEmptyRows();
                case InsertImage.Operation _: return new InsertImage();
                case InsertTable.Operation _: return new InsertTable();
                case InsertText.Operation _: return new InsertText();
                default: throw new Exception($"Не найден исполнитель для операции {operation?.GetType().Name}");
            }
        }
    }
}