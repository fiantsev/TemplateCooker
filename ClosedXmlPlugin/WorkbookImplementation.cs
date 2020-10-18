using ClosedXML.Excel;
using PluginAbstraction;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClosedXmlPlugin
{
    public class WorkbookImplementation : IWorkbookAbstraction
    {
        private XLWorkbook _workbook;
        private bool _recalculateFormulasOnSave;

        public WorkbookImplementation()
        {
            _workbook = new XLWorkbook();
        }

        public WorkbookImplementation(Stream stream)
        {
            _workbook = new XLWorkbook(stream);
        }

        public WorkbookImplementation(XLWorkbook workbook)
        {
            _workbook = workbook;
        }

        public ISheetAbstraction GetSheet(int index)
        {
            var sheet = new SheetImplementation(_workbook.Worksheet(index + 1));
            return sheet;
        }

        public IEnumerable<ISheetAbstraction> GetSheets()
        {
            var sheets = _workbook.Worksheets.Select(x => new SheetImplementation(x));
            return sheets;
        }

        public void AddPicture(MemoryStream imageStream, int sheetIndex, int rowIndex, int columnIndex)
        {
            var sheet = _workbook.Worksheet(sheetIndex + 1);

            sheet
                .AddPicture(imageStream)
                .MoveTo(sheet.Cell(rowIndex + 1, columnIndex + 1));
        }

        public void SetProperties(bool forceFullCalculation, bool fullCalculationOnLoad, bool recalculateFormulasOnSave)
        {
            _workbook.ForceFullCalculation = forceFullCalculation;
            _workbook.FullCalculationOnLoad = fullCalculationOnLoad;
            _recalculateFormulasOnSave = recalculateFormulasOnSave;
        }

        public void Save(Stream stream)
        {
            _workbook.SaveAs(stream, false, _recalculateFormulasOnSave);
        }

        public void Dispose()
        {
            _workbook.Dispose();
        }
    }
}