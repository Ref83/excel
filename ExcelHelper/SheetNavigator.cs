using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelHelper
{
    /// <summary>
    /// Навигатор копии страницы Excel
    /// </summary>
    public class SheetNavigator
    {
        #region Stratic Fields
        private static readonly string maxColumnAlias = "XFD";
        private static readonly IDictionary<string, int> columnAliases = new Dictionary<string, int>();
        #endregion

        #region Fields
        private object[,] cells;
        private IList<Row> rows = new List<Row>();
        private IList<Column> columns = new List<Column>();
        #endregion

        #region Constructors
        static SheetNavigator()
        {
            var maxColumns = ExcelNavigationHelper.GetColumnNumberFor(SheetNavigator.maxColumnAlias);
            for (int i = 1; i <= maxColumns; i++)
            {
                SheetNavigator.columnAliases.Add(ExcelNavigationHelper.ConvertNumberToColumn(i), i);
            }
        }

        internal SheetNavigator(object[,] cells)
        {
            if (cells == null)
                throw new ArgumentNullException("cells");

            this.cells = cells;
            for (int i = 1; i < cells.GetLength(0); i++)
            {
                this.rows.Add(new Row(i, this));
            }

            for (int i = 1; i < cells.GetLength(1); i++)
            {
                this.columns.Add(new Column(ExcelNavigationHelper.ConvertNumberToColumn(i), this));
            }

        }
        #endregion

        #region Helpers

        #endregion

        #region Public Interface
        /// <summary>
        /// Получение значения ячейки по указаным координатам
        /// </summary>
        /// <typeparam name="T">Тип к которому приводится значение ячейки</typeparam>
        /// <param name="column">Псевдоним колонки. Указывается заглавными буквами в диапазоне от A до XFD</param>
        /// <param name="row">Номер строки. Больше или равен нулю</param>
        /// <returns>Значение в ячейке приведенное к типу T</returns>
        public T GetValueFromCell<T>(string column, int row)
        {
            if (0 >= row || row >= this.cells.GetLength(0))
                throw new ArgumentOutOfRangeException("Параметр row должен быть больше нуля и меньше количества строк в управляемом массиве");

            if (!SheetNavigator.columnAliases.ContainsKey(column))
                throw new ArgumentOutOfRangeException(String.Format("Параметр column должен быть в диапазоне A..{0} и задан только заглавными буквами", SheetNavigator.maxColumnAlias));

            var columnNumber = SheetNavigator.columnAliases[column];
            if (columnNumber >= this.cells.GetLength(1))
                throw new ArgumentOutOfRangeException("Параметр column должен быть меньше количества колонок в управляемомо массиве");

            try
            {
                var cell = cells[row, columnNumber];
                return (T)cell;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(String.Format("Ошибка получения данных из ячейки {0}{1}", column, row.ToString()), ex);
            }
        }

        /// <summary>
        /// Получение значения ячейки по наименованию ячейки
        /// </summary>
        /// <typeparam name="T">Тип к которому приводится значение ячейки</typeparam>
        /// <param name="cellName">Наименование ячейки. Состоит из псевдонима столбца и номера строки. Например "A1" или "ZZ56"</param>
        /// <returns>Значение в ячейке приведенное к типу T</returns>
        public T GetValueFromCell<T>(string cellName)
        {
            string column;
            int row;

            try
            {
                column = ExcelNavigationHelper.GetColumnFromCellName(cellName);
                row = ExcelNavigationHelper.GetRowNumberFor(cellName);
            }
            catch (Exception ex)
            {
                throw new ArgumentOutOfRangeException(String.Format("Ошибка получения строки или колонки из адреса ячейки {0}", cellName), ex);
            }

            return GetValueFromCell<T>(column, row);
        }

        /// <summary>
        /// Проверка возможности приведения значения ячейки к типу T
        /// </summary>
        /// <typeparam name="T">Тип к которому приводится значение ячейки</typeparam>
        /// <param name="column">Псевдоним колонки. Указывается заглавными буквами в диапазоне от A до XFD</param>
        /// <param name="row">Номер строки. Больше или равен нулю</param>
        /// <returns>true - если значение можно привести к типу T, иначе false</returns>
        public bool CheckDataTypeInCell<T>(string column, int row)
        {
            var isTypeMatch = true;
            try
            {
                var value = GetValueFromCell<T>(column, row);
            }
            catch
            {
                isTypeMatch = false;
            }

            return isTypeMatch;
        }

        /// <summary>
        /// Проверка значения указаной ячейки на null
        /// </summary>
        /// <param name="column">Псевдоним колонки. Указывается заглавными буквами в диапазоне от A до XFD</param>
        /// <param name="row">Номер строки. Больше или равен нулю</param>
        /// <returns>true - если значение null, иначе false</returns>
        public bool IsNullInCell(string column, int row)
        {
            var value = GetValueFromCell<object>(column, row);
            return value == null;
        }

        /// <summary>
        /// Перечисление строк
        /// </summary>
        public IEnumerable<Row> Rows
        {
            get { return this.rows; }
        }

        /// <summary>
        /// Перечисление столбцов
        /// </summary>
        public IEnumerable<Column> Columns
        {
            get { return this.columns; }
        }
        #endregion
    }
}
