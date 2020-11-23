﻿using ClosedXmlPlugin;
using PluginAbstraction;
using System.Collections.Generic;
using System.IO;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Recipes;

namespace TemplateCooker.Service.Cooker
{
    public class TemplateCooker
    {
        private IWorkbookAbstraction _workbook;

        public TemplateCooker(Stream workbookStream)
        {
            workbookStream.Position = 0;
            var plugin = new ClosedXmlPluginImplementation();
            _workbook = plugin.OpenWorkbook(workbookStream);
        }

        public List<Marker> ExtractMarkers(ExtractMarkersRecipe.Options options)
        {
            options.Workbook = _workbook;
            var markers = new ExtractMarkersRecipe(options).Cook();
            return markers;
        }

        public TemplateCooker InjectData(InjectRecipe.Options options)
        {
            options.Workbook = _workbook;
            new InjectRecipe(options).Cook();
            return this;
        }

        public TemplateCooker SetCustomProperties(SetCustomPropertiesRecipe.Options options)
        {
            options.Workbook = _workbook;
            new SetCustomPropertiesRecipe(options).Cook();
            return this;
        }

        public MemoryStream Build()
        {
            var resultStream = new MemoryStream();
            _workbook.Save(resultStream);
            resultStream.Position = 0;

            //делаем инстанс более не юзабельным
            _workbook = null;

            return resultStream;
        }
    }
}