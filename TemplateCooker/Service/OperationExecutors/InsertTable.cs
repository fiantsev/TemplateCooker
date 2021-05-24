using PluginAbstraction;
using System.Collections.Generic;
using System.Linq;
using TemplateCooking.Domain.Layout;
using TemplateCooking.Domain.ResourceObjects;

namespace TemplateCooking.Service.OperationExecutors
{
    /// <summary>
    /// Вставить табличные данные в указанную ячейку
    /// </summary>
    public class InsertTable : IOperationExecutor
    {
        public class Operation : AbstractOperation
        {
            /// <summary>
            /// позиция верхней левой ячейки начиная с которой будет располагаться таблица
            /// </summary>
            public SrcPosition Position { get; set; }
            /// <summary>
            /// данные таблицы
            /// </summary>
            public TableWithHeaders Table { get; set; }
            /// <summary>
            /// Если больше нуля - то включаются фиксированные отступы по строкам и не учитывается высота смердженных ячеек
            /// </summary>
            public int FixedRowStep { get; set; }
            /// <summary>
            /// Если больше нуля - то включаются фиксированные отступы по столбцам и не учитывается ширина смердженных ячеек
            /// </summary>
            public int FixedColumnStep { get; set; }
            /// <summary>
            /// Объединять одинаковые ячейки в заголовках столбцов
            /// </summary>
            public bool MergeColumnHeaders { get; set; }
            /// <summary>
            /// Объединять одинаковые ячейки в заголовках строк
            /// </summary>
            public bool MergeRowHeaders { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Operation)untypedOptions;
            var sheet = workbook.GetSheet(options.Position.SheetIndex);
            var topLeftCell = sheet.GetRow(options.Position.RowIndex).GetCell(options.Position.ColumnIndex);

            //удаляем маркер
            topLeftCell.SetValue(string.Empty);

            //вставляем заголовки столбцов
            var rowHeadersWidth = options.Table.RowHeaders.FirstOrDefault()?.Count ?? 0;
            var columnHeadersTopLeftCell = topLeftCell.GetMergedCells(1, rowHeadersWidth + 1, options.FixedRowStep, options.FixedColumnStep).LastOrDefault();
            if (options.MergeColumnHeaders)
                InsertHierarchyTable(columnHeadersTopLeftCell, options.FixedRowStep, options.FixedColumnStep, options.Table.ColumnHeaders, MergeDirection.Horizontal);
            else
                InsertFlatTable(columnHeadersTopLeftCell, options.FixedRowStep, options.FixedColumnStep, options.Table.ColumnHeaders);

            //вставляем заголовки строк
            var columnHeadersHeight = options.Table.ColumnHeaders.Count;
            var rowHeadersTopLeftCell = topLeftCell.GetMergedCells(columnHeadersHeight + 1, 1, options.FixedRowStep, options.FixedColumnStep).LastOrDefault();
            if (options.MergeRowHeaders)
                InsertHierarchyTable(rowHeadersTopLeftCell, options.FixedRowStep, options.FixedColumnStep, options.Table.RowHeaders, MergeDirection.Vertical);
            else
                InsertFlatTable(rowHeadersTopLeftCell, options.FixedRowStep, options.FixedColumnStep, options.Table.RowHeaders);

            //вставляем заголовки строк
            var bodyTopLeftCell = topLeftCell.GetMergedCells(columnHeadersHeight + 1, rowHeadersWidth + 1, options.FixedRowStep, options.FixedColumnStep).LastOrDefault();
            InsertFlatTable(bodyTopLeftCell, options.FixedRowStep, options.FixedColumnStep, options.Table.Body);

            //кросс заголовок (верхняя левая ячейка) - мерджим в одну большую ячейку
            if (columnHeadersHeight > 0 && rowHeadersWidth > 0)
            {
                var bottomCorner = topLeftCell.GetMergedCells(columnHeadersHeight, rowHeadersWidth, options.FixedRowStep, options.FixedColumnStep).LastOrDefault().GetMergedRange().BottomRightCell();
                var crossHeaderRange = sheet.GetRange(topLeftCell, bottomCorner);
                crossHeaderRange.Merge();
            }
        }

        /// <summary>
        /// вставляет обычную таблицу
        /// </summary>
        private void InsertFlatTable(ICellAbstraction topLeftCell, int fixedRowStep, int fixedColumnStep, List<List<object>> table)
        {
            var rowCount = table.Count;
            var columnCount = rowCount == 0 ? 0 : table[0].Count;

            //здесь мы вставляем контент в ячейки
            var cellsToInsert = table.Select((row, rowIndex) => row.Select((cellValue, cellIndex) => new CellInfo { CellValue = cellValue, Cell = null }).ToList()).ToList();
            var cellCounter = 0;
            foreach (var cell in topLeftCell.GetMergedCells(rowCount, columnCount, fixedRowStep, fixedColumnStep))
            {
                var rowIndex = cellCounter / columnCount;
                var columnIndex = cellCounter % columnCount;
                cellsToInsert[rowIndex][columnIndex].Cell = cell;
                var value = table[rowIndex][columnIndex];
                cell.SetValue(value);
                ++cellCounter;
            }
        }

        /// <summary>
        /// вставляет таблицу с иерархией в заголовках
        /// </summary>
        private void InsertHierarchyTable(ICellAbstraction topLeftCell, int fixedRowStep, int fixedColumnStep, List<List<object>> table, MergeDirection mergeDirection)
        {
            var rowCount = table.Count;
            var columnCount = rowCount == 0 ? 0 : table[0].Count;

            //заполняем таблицу cellInfos
            var cellInfos = table.Select((row, rowIndex) => row.Select((cellValue, cellIndex) => new CellInfo { CellValue = cellValue, Cell = null }).ToList()).ToList();
            var cellCounter = 0;
            foreach (var cell in topLeftCell.GetMergedCells(rowCount, columnCount, fixedRowStep, fixedColumnStep))
            {
                var rowIndex = cellCounter / columnCount;
                var columnIndex = cellCounter % columnCount;
                cellInfos[rowIndex][columnIndex].Cell = cell;
                ++cellCounter;
            }

            //создаем корень дерева и выращиваем дерево на основе пришедших данных в cellInfos
            var treeRoot = new TreeNode<MergeGroup>(null);
            //транспонируем в случае когда нужно расчитать для заголовков строк регионы объединения
            var cells = mergeDirection == MergeDirection.Vertical ? cellInfos.Transpose() : cellInfos;
            CreateTree(treeRoot, cells);

            //отрисовываем в экселе дерево
            treeRoot.Children //сразу начинаем с детей потому что рутовый элемент является вспомогательным и ничего не содержит
                .ForEach(child => child.Traverse(node =>
                {
                    if (node.Data.Cells.Count > 1)//если количество ячеек в группе больше одной - то необходимо мерджить ячейки
                    {
                        var first = node.Data.Cells.First();
                        var last = node.Data.Cells.Last();
                        var range = first.Sheet.GetRange(first, last.GetMergedRange().BottomRightCell());
                        range.Merge();
                        first.SetValue(node.Data.Value);
                    }
                    else //в остальных случаях просто берем первую ячейку и выводим ее значение в эксель
                    {
                        var first = node.Data.Cells.First();
                        first.SetValue(node.Data.Value);
                    }

                    return true;
                }));
        }

        /// <summary>
        /// превращает плоскую матрицу значений в иерархичное представление ввиде дерева.
        /// Алгоритм: соседние ячейки с одинаковым значением сливаются в одну, при условие что на верхнем уровне принадлежат к одному родителю
        /// </summary>
        /// <param name="parentNode"></param>
        /// <param name="cellInfos"></param>
        private void CreateTree(TreeNode<MergeGroup> parentNode, List<List<CellInfo>> cellInfos)
        {
            if (cellInfos.Count == 0)
                return;

            var firstRow = cellInfos.FirstOrDefault();
            var groups = new List<MergeGroup>
            {
                new MergeGroup(firstRow.FirstOrDefault()?.CellValue) { StartIndex = 0, EndIndex = 0 }
            };
            var currentGroup = groups[0];

            for (var i = 0; i < firstRow.Count; ++i)
            {
                var currentCellInfo = firstRow[i];
                if (currentCellInfo.CellValue.Equals(currentGroup.Value))
                {
                    currentGroup.Cells.Add(currentCellInfo.Cell);
                    currentGroup.EndIndex = i;
                }
                else
                {
                    currentGroup = new MergeGroup(currentCellInfo.CellValue)
                    {
                        Cells = new List<ICellAbstraction> { currentCellInfo.Cell },
                        StartIndex = i,
                        EndIndex = i
                    };
                    groups.Add(currentGroup);
                }
            }

            groups
                .Where(x => x.Cells.Count > 0)
                .ToList()
                .ForEach(group =>
                {
                    var nodeForGroup = parentNode.AddChild(group);
                    var subset = cellInfos.Skip(1)
                        .Select(x => x.Skip(group.StartIndex).Take(group.Cells.Count).ToList()).ToList();
                    CreateTree(nodeForGroup, subset);
                });
        }

    }

    class CellInfo
    {
        public object CellValue { get; set; }
        public ICellAbstraction Cell { get; set; }
    }

    enum MergeDirection
    {
        Horizontal,
        Vertical,
    }

    class MergeGroup
    {
        public object Value { get; set; }
        public List<ICellAbstraction> Cells { get; set; }

        public int StartIndex { get; set; }
        public int EndIndex { get; set; }

        public MergeGroup(object value)
        {
            Value = value;
            Cells = new List<ICellAbstraction>();
        }
    }
};