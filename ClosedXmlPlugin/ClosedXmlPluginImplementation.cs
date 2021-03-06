﻿using PluginAbstraction;
using System.IO;

namespace ClosedXmlPlugin
{
    public class ClosedXmlPluginImplementation : IPluginAbstraction
    {
        public IWorkbookAbstraction CreateEmptyWorkbook()
        {
            return new WorkbookImplementation();
        }

        public IWorkbookAbstraction OpenWorkbook(Stream stream)
        {
            return new WorkbookImplementation(stream);
        }
    }
}