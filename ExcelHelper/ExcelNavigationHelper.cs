using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelHelper
{
    internal static class ExcelNavigationHelper
    {
        private static string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        #region Helpers
        private static int GetRowPositionFor(string cellName)
        {
            var index = 1;

            foreach (char letter in cellName)
            {
                if (Char.IsDigit(letter))
                    break;
                index++;
            }

            return index - 1;
        }

        private static string GetRowFromCellName(string cellName)
        {
            int rowPosition = GetRowPositionFor(cellName);
            return cellName.Substring(rowPosition, cellName.Length - rowPosition);
        }
        #endregion

        #region Public Interface

        public static int ConvertColumnToNumber(string column)
        {
            int index = 0;
            int power = 1;
            foreach (char simbol in column.Reverse())
            {
                index = index + power * (alphabet.IndexOf(simbol) + 1);
                power = power * alphabet.Length;
            }
            return index;
        }

        public static string ConvertNumberToColumn(int number)
        {
            int index = number;
            string result = "";
            while (index > 0)
            {
                result = alphabet[index % alphabet.Length == 0 ? alphabet.Length - 1 : index % alphabet.Length - 1] + result;
                index = index / alphabet.Length - (index % alphabet.Length == 0 ? 1 : 0);
            };
            return result;
        }

        public static string GetNextColumn(string current)
        {
            int index = ConvertColumnToNumber(current);
            index++;
            return ConvertNumberToColumn(index);
        }

        public static string GetPrevColumn(string current)
        {
            int index = ConvertColumnToNumber(current);
            index--;
            return ConvertNumberToColumn(index);
        }

        public static string ShiftColumn(string current, int shift)
        {
            int index = ConvertColumnToNumber(current);
            index = index + shift;
            return ConvertNumberToColumn(index);
        }

        public static string GetColumnFromCellName(string cellName)
        {
            return cellName.Substring(0, GetRowPositionFor(cellName));
        }

        public static int GetColumnNumberFor(string cellName)
        {
            var column = GetColumnFromCellName(cellName);
            return ConvertColumnToNumber(column);
        }

        public static int GetRowNumberFor(string cellName)
        {
            var row = GetRowFromCellName(cellName);
            int result;
            if (!Int32.TryParse(row, out result))
                throw new InvalidOperationException(String.Format("Не удалось получить строку из адреса ячейки {0}", cellName));
            return result;
        }

        #endregion
    }
}
