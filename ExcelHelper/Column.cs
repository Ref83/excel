using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelHelper
{
    public class Column
    {
        #region Fields
        private string columnAlias;
        private SheetNavigator navigator;
        #endregion

        #region Constructors
        internal Column(string columnAlias, SheetNavigator navigator)
        {
            this.columnAlias = columnAlias;
            this.navigator = navigator;
        }
        #endregion

        #region Public Interface
        public string ColumnAlias
        {
            get { return this.columnAlias; }
        }

        /// <summary>
        /// Получение значения ячейки в колонке
        /// </summary>
        /// <typeparam name="T">Тип к которому приводится значение ячейки</typeparam>
        /// <param name="row">Номер строки. Больше или равен нулю</param>
        /// <returns>Значение в ячейке приведенное к типу T</returns>
        public T GetValueFromRow<T>(int row)
        {
            return this.navigator.GetValueFromCell<T>(this.columnAlias, row);
        }

        /// <summary>
        /// Проверка возможности приведения значения ячейки в строке к типу T
        /// </summary>
        /// <typeparam name="T">Тип к которому приводится значение ячейки</typeparam>
        /// <param name="row">Номер строки. Больше или равен нулю</param>
        /// <returns>true - если значение можно привести к типу T, иначе false</returns>
        public bool CheckDataTypeInRow<T>(int row)
        {
            return this.navigator.CheckDataTypeInCell<T>(this.columnAlias, row);
        }

        /// <summary>
        /// Проверка значения указаной ячейки в колонке на null
        /// </summary>
        /// <param name="row">Номер строки. Больше или равен нулю</param>
        /// <returns>true - если значение null, иначе false</returns>
        public bool IsNullInRow(int row)
        {
            return this.navigator.IsNullInCell(this.columnAlias, row);
        }

        #endregion
    }
}
