using NPOI.SS.UserModel;
using PluginAbstraction;
using System.Collections.Generic;
using System.Diagnostics;

namespace NpoiPlugin
{
    [DebuggerDisplay("{_sheet}")]
    public class SheetImplementation : ISheetAbstraction
    {
        private ISheet _sheet;

        public SheetImplementation(ISheet sheet)
        {
            _sheet = sheet;
        }

        public int SheetIndex => _sheet.Workbook.GetSheetIndex(_sheet);
        public string SheetName => _sheet.SheetName;

        public IRowAbstraction GetRow(int index)
        {
            var row = _sheet.GetRow(index) ?? _sheet.CreateRow(index);
            return new RowImplementation(row);
        }

        public IEnumerable<IRowAbstraction> GetRows()
        {
            var enumerator = _sheet.GetEnumerator();
            enumerator.MoveNext();
            while (enumerator.Current != null)
            {
                yield return new RowImplementation((IRow)enumerator.Current);
                enumerator.MoveNext();
            }
        }

        public IEnumerable<IRowAbstraction> GetUsedRows()
        {
            for (var rowIndex = _sheet.FirstRowNum; rowIndex <= _sheet.LastRowNum; ++rowIndex)
            {
                var row = _sheet.GetRow(rowIndex);
                if (row == null) continue;
                yield return new RowImplementation(row);
            }
        }

        public void Dispose()
        {
        }

        public IRangeAbstraction GetRange(ICellAbstraction from, ICellAbstraction to)
        {
            throw new System.NotImplementedException();
        }
    }
}