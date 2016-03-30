using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelHelper
{
    /// <summary>
    /// Строка
    /// </summary>
    public class Row
    {
        #region Fields
        private int rowNumber;
        private SheetNavigator navigator;
        #endregion

        #region Constructors
        internal Row(int rowNumber, SheetNavigator navigator)
        {
            this.rowNumber = rowNumber;
            this.navigator = navigator;
        }
        #endregion

        #region Public Interface
        /// <summary>
        /// Номер строки
        /// </summary>
        public int RowNumber
        {
            get { return this.rowNumber; }
        }

        /// <summary>
        /// Получение значения ячейки в задвной колонке в текущей строке
        /// </summary>
        /// <typeparam name="T">Тип к которому приводится значение ячейки</typeparam>
        /// <param name="columnAlias">Псевдоним колонки. Указывается заглавными буквами в диапазоне от A до XFD</param>
        /// <returns>Значение в ячейке приведенное к типу T</returns>
        public T GetValueFromCollumn<T>(string columnAlias)
        {
            return this.navigator.GetValueFromCell<T>(columnAlias, this.rowNumber);
        }

        /// <summary>
        /// Проверка возможности приведения значения ячейки в строке к типу T
        /// </summary>
        /// <typeparam name="T">Тип к которому приводится значение ячейки</typeparam>
        /// <param name="columnAlias">Псевдоним колонки. Указывается заглавными буквами в диапазоне от A до XFD</param>
        /// <returns>true - если значение можно привести к типу T, иначе false</returns>
        public bool CheckDataTypeInCollumn<T>(string columnAlias)
        {
            return this.navigator.CheckDataTypeInCell<T>(columnAlias, this.rowNumber);
        }

        /// <summary>
        /// Проверка значения указаной ячейки в строке на null
        /// </summary>
        /// <param name="columnAlias">Псевдоним колонки. Указывается заглавными буквами в диапазоне от A до XFD</param>
        /// <returns>true - если значение null, иначе false</returns>
        public bool IsNullInColumn(string columnAlias)
        {
            return this.navigator.IsNullInCell(columnAlias, this.rowNumber);
        }

        #endregion
    }
}
