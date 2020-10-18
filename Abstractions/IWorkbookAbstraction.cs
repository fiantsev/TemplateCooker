using System;
using System.Collections.Generic;
using System.IO;

namespace PluginAbstraction
{
    public interface IWorkbookAbstraction : IDisposable
    {
        IEnumerable<ISheetAbstraction> GetSheets();
        ISheetAbstraction GetSheet(int index);
        void AddPicture(MemoryStream imageStream, int sheetIndex, int rowIndex, int columnIndex);

        void SetProperties(bool forceFullCalculation, bool fullCalculationOnLoad, bool recalculateFormulasOnSave);
        void Save(Stream stream);
    }
}