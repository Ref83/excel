using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelHelper;

namespace ExcelHelper.Tests
{
    [TestClass]
    public class SheetNavigatorTests
    {
        #region Fields
        private SheetNavigator navigator;
        #endregion

        #region Helpers
        private void CheckException(Action action, Type exceptionType)
        {
            try
            {
                action();
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsInstanceOfType(ex, exceptionType);
            }
        }
        #endregion

        #region Fixture
        [TestInitialize]
        public void Initialize()
        {
            this.navigator = new SheetNavigator(new object[,] { { null, null, null }, { null, "A1", "B1" }, { null, "A2", "B2" } });
        }
        #endregion

        #region Tests for data in cells
        [TestMethod]
        public void Can_Get_Cell_By_Column_And_Row()
        {
            Assert.AreEqual("A1", this.navigator.GetValueFromCell<string>("A", 1));
            Assert.AreEqual("A2", this.navigator.GetValueFromCell<string>("A", 2));
            Assert.AreEqual("B1", this.navigator.GetValueFromCell<string>("B", 1));
            Assert.AreEqual("B2", this.navigator.GetValueFromCell<string>("B", 2));
        }

        [TestMethod]
        public void Can_Get_Cell_By_CellName()
        {
            Assert.AreEqual("B2", this.navigator.GetValueFromCell<string>("B2"));
        }

        [TestMethod]
        public void Can_Check_Data_Type_In_Cell()
        {
            Assert.IsFalse(this.navigator.CheckDataTypeInCell<DateTime>("A", 1));
            Assert.IsFalse(this.navigator.CheckDataTypeInCell<int>("A", 1));
            Assert.IsFalse(this.navigator.CheckDataTypeInCell<decimal?>("A", 1));
            Assert.IsTrue(this.navigator.CheckDataTypeInCell<string>("A", 1));
            Assert.IsTrue(this.navigator.CheckDataTypeInCell<object>("A", 1));
        }

        [TestMethod]
        public void Can_Check_Null_In_Cell()
        {
            var navigatorWithNulls = new SheetNavigator(new object[,] { { null, null, null }, { null, null, "B1" }, { null, "A2", "B2" } });

            Assert.IsTrue(navigatorWithNulls.IsNullInCell("A", 1));
            Assert.IsFalse(navigatorWithNulls.IsNullInCell("A", 2));
        }

        [TestMethod]
        public void Can_Complex_Navigation()
        {
            DateTime baseDate = new DateTime(2015, 12, 31);
            decimal baseValueColumn = 10M;

            var complexNavigator = new SheetNavigator(new object[,] { { null, null, null, null, null, null }, { null, new DateTime(baseDate.Year, 12, 31), "", new DateTime(baseDate.Year + 1, 12, 31), "", new DateTime(baseDate.Year + 2, 12, 31) }, { null, baseValueColumn, "", baseValueColumn, "", baseValueColumn }, { null, 2 * baseValueColumn, "", 2 * baseValueColumn, "", 2 * baseValueColumn }, { null, 3 * baseValueColumn, "", 3 * baseValueColumn, "", 3 * baseValueColumn } });

            Assert.AreEqual(3, complexNavigator.Columns.Where(c => c.CheckDataTypeInRow<DateTime>(1)).Count());

            // Обходим колонки где есть дата
            foreach (var column in complexNavigator.Columns.Where(c => c.CheckDataTypeInRow<DateTime>(1)))
            {
                Assert.AreEqual(baseDate, column.GetValueFromRow<DateTime>(1));

                // Обходим строки с номером >= 3
                foreach (var row in complexNavigator.Rows.Where(r => r.RowNumber >= 3))
                {
                    Assert.AreEqual(baseValueColumn * (row.RowNumber - 1), column.GetValueFromRow<decimal>(row.RowNumber));
                }

                baseDate = new DateTime(baseDate.Year + 1, 12, 31);
            }
        }
        #endregion

        #region Tests enums
        [TestMethod]
        public void Can_Enum_Rows()
        {
            Assert.AreEqual(2, this.navigator.Rows.Count());

            foreach (var row in this.navigator.Rows)
            {
                Assert.AreEqual(this.navigator.GetValueFromCell<string>("A", row.RowNumber), row.GetValueFromCollumn<string>("A"));
                Assert.AreEqual(this.navigator.GetValueFromCell<string>("B", row.RowNumber), row.GetValueFromCollumn<string>("B"));
            }
        }

        [TestMethod]
        public void Can_Enum_Columns()
        {
            Assert.AreEqual(2, this.navigator.Columns.Count());

            foreach (var column in this.navigator.Columns)
            {
                Assert.AreEqual(this.navigator.GetValueFromCell<string>(column.ColumnAlias, 1), column.GetValueFromRow<string>(1));
                Assert.AreEqual(this.navigator.GetValueFromCell<string>(column.ColumnAlias, 2), column.GetValueFromRow<string>(2));
            }
        }

        [TestMethod]
        public void Can_Filter_Enum_Rows()
        {
            Assert.AreEqual(1, this.navigator.Rows.Where(r => r.GetValueFromCollumn<string>("A").Contains("1")).Count());
        }

        [TestMethod]
        public void Can_Filter_Enum_Columns()
        {
            Assert.AreEqual(1, this.navigator.Columns.Where(c => c.GetValueFromRow<string>(1).Contains("A")).Count());
        }


        #endregion

        #region Tests for check errors
        [TestMethod]
        public void Exception_When_Create_With_Null_Argument()
        {
            CheckException(() => new SheetNavigator(null), typeof(ArgumentNullException));
        }

        [TestMethod]
        public void Exception_When_Try_Get_Cell_With_Negative_Or_Zerro_Row()
        {
            CheckException(() => this.navigator.GetValueFromCell<string>("A", -1), typeof(ArgumentOutOfRangeException));
            CheckException(() => this.navigator.GetValueFromCell<string>("A", 0), typeof(ArgumentOutOfRangeException));
        }

        [TestMethod]
        public void Exception_When_Try_Get_Cell_With_Big_Row_Number()
        {
            CheckException(() => this.navigator.GetValueFromCell<string>("A", 3), typeof(ArgumentOutOfRangeException));
        }

        [TestMethod]
        public void Exception_When_Try_Get_Cell_With_Big_ColumnAlias()
        {
            CheckException(() => this.navigator.GetValueFromCell<string>("Z", 2), typeof(ArgumentOutOfRangeException));
        }

        [TestMethod]
        public void Exception_When_Try_Get_Cell_With_Not_Exists_Column()
        {
            CheckException(() => this.navigator.GetValueFromCell<string>("not exists column alias", 1), typeof(ArgumentOutOfRangeException));
        }

        [TestMethod]
        public void Exception_When_Try_Get_Cell_With_ColumnAlias_By_Small_Latters()
        {
            CheckException(() => this.navigator.GetValueFromCell<string>("a", 1), typeof(ArgumentOutOfRangeException));
        }

        [TestMethod]
        public void Exception_When_Try_Get_Cell_By_CellName_With_Not_Exists_Column()
        {
            CheckException(() => this.navigator.GetValueFromCell<string>("ZZZZZ2"), typeof(ArgumentOutOfRangeException));
        }

        [TestMethod]
        public void Exception_When_Try_Get_Cell_By_CellName_With_Negative_Or_Zerro_Row()
        {
            CheckException(() => this.navigator.GetValueFromCell<string>("A-2"), typeof(ArgumentOutOfRangeException));
            CheckException(() => this.navigator.GetValueFromCell<string>("A0"), typeof(ArgumentOutOfRangeException));
        }

        [TestMethod]
        public void Exception_When_Try_Get_Cell_By_CellName_Without_Row()
        {
            CheckException(() => this.navigator.GetValueFromCell<string>("A"), typeof(ArgumentOutOfRangeException));
        }

        [TestMethod]
        public void Exception_When_Try_Get_Cell_By_CellName_With_Total_Bad_Name()
        {
            CheckException(() => this.navigator.GetValueFromCell<string>("bad cell name"), typeof(ArgumentOutOfRangeException));
            CheckException(() => this.navigator.GetValueFromCell<string>("a1"), typeof(ArgumentOutOfRangeException));
        }
        #endregion

    }
}
