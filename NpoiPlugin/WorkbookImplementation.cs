using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PluginAbstraction;
using PluginAbstraction.Exceptions;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace NpoiPlugin
{
    public class WorkbookImplementation : IWorkbookAbstraction
    {
        private IWorkbook _workbook;
        private bool _recalculateFormulasOnSave;

        public WorkbookImplementation()
        {
            _workbook = new XSSFWorkbook();
        }

        public WorkbookImplementation(Stream stream)
        {
            _workbook = new XSSFWorkbook(stream);
        }

        public WorkbookImplementation(IWorkbook workbook)
        {
            _workbook = workbook;
        }

        public ISheetAbstraction GetSheet(int index)
        {
            var sheet = new SheetImplementation(_workbook.GetSheetAt(index));
            return sheet;
        }

        public IEnumerable<ISheetAbstraction> GetSheets()
        {
            foreach (var sheetIndex in Enumerable.Range(0, _workbook.NumberOfSheets))
            {
                var sheet = _workbook.GetSheetAt(sheetIndex);
                yield return new SheetImplementation(sheet);
            }
            //var sheets = _workbook.Sheets .Worksheets.Select(x => new SheetImplementation(x));
            //return sheets;
        }

        public void AddPicture(MemoryStream imageStream, int sheetIndex, int rowIndex, int columnIndex)
        {
            throw new PluginUnsupportedException("AddPicture");
            //var sheet = _workbook.Worksheet(sheetIndex + 1);

            //sheet
            //    .AddPicture(imageStream)
            //    .MoveTo(sheet.Cell(rowIndex + 1, columnIndex + 1));
        }

        public void SetCustomProperties(CustomProperties customProperties)
        {
            //throw new PluginUnsupportedException();
            //_workbook.ForceFullCalculation = forceFullCalculation;
            //_workbook.FullCalculationOnLoad = fullCalculationOnLoad;
            //_recalculateFormulasOnSave = recalculateFormulasOnSave;
        }

        public void Save(Stream stream)
        {
            _workbook.Write(stream);
        }

        public void Dispose()
        {
            _workbook.Close();
        }

        public ISheetAbstraction AddSheet(string name)
        {
            return new SheetImplementation(_workbook.CreateSheet(name));
        }
    }
}