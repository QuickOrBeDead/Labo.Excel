using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using Labo.Common.Exceptions;
using Labo.Common.Reflection;
using Labo.Common.Utils;
using Microsoft.Office.Interop.Excel;

namespace Labo.Excel
{
    /// <summary>
    /// The office interop excel reader writer class.
    /// </summary>
    public sealed class OfficeInteropExcelReaderWriter : BaseExcelReaderWriter
    {
        /// <summary>
        /// Reads the specified excel internal method.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="mapper">The mapper.</param>
        /// <param name="onStartReading">The on start reading.</param>
        /// <returns></returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling")]
        public override List<T> ReadInternal<T>(string fileName, Func<ExcelRowReadOperation, T> mapper, Action<ExcelReadStartOperationInfo> onStartReading = null)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentNullException(nameof(fileName));
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));

            Application xlApp = null;
            Workbook xlWorkBook = null;
            Worksheet xlWorkSheet = null;
            object misValue = Missing.Value;

            try
            {

                xlApp = new Application();

                if (xlApp == null)
                {
                    throw new CoreLevelException("EXCEL could not be started. Check that your office installation and project references are correct.");
                }

                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, string.Empty, string.Empty, true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

                Range excelRange = xlWorkSheet.UsedRange;
                object[,] values = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
                int rowsLength = values.GetUpperBound(0);
                int columnsLength = values.GetUpperBound(1);
                List<T> result = new List<T>();

                ExcelDictionaryColumns excelDictionaryColumns = new ExcelDictionaryColumns();
                string[] columns = new string[columnsLength + 1];
                for (int i = 1; i < columns.Length; i++)
                {
                    string columnName = ConvertUtils.ChangeType<string>(values.GetValue(1, i));
                    columns[i] = columnName;
                    excelDictionaryColumns.Add(columnName);
                }

                excelDictionaryColumns.EnsureColumnsNotNull();
                excelDictionaryColumns.EnsureColumnsAreUnique();

                int lastRowNum = rowsLength;
                if (onStartReading != null)
                {
                    onStartReading(new ExcelReadStartOperationInfo(lastRowNum, columns));
                }
                for (int i = 2; i <= rowsLength; i++)
                {
                    ExcelDictionary row = new ExcelDictionary(excelDictionaryColumns);
                    for (int j = 1; j <= columnsLength; j++)
                    {
                        object value = values.GetValue(i, j);

                        row[columns[j]] = value;
                    }

                    ExcelRowReadOperation readOperation = new ExcelRowReadOperation(row);
                    T item = mapper(readOperation);
                    if (readOperation.Cancel)
                    {
                        break;
                    }

                    result.Add(item);
                }
                return result;
            }
            finally 
            {
                if (xlWorkBook != null)
                {
                    xlWorkBook.Close(misValue, misValue, misValue);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                }

                ReleaseObject(xlWorkSheet);
                ReleaseObject(xlWorkBook);
                ReleaseObject(xlApp);
            }
        }

        /// <summary>
        /// Writes the specified table to the excel file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="list">The list.</param>
        /// <param name="autoResizeColumns">if set to <c>true</c> [automatic resize columns].</param>
        /// <param name="culture">The culture.</param>
        public override void Write<T>(string fileName, List<T> list, bool autoResizeColumns = false, CultureInfo culture = null)
        {
            ExcelTable table = ConvertListToExcelTableObject(list);

            Write(fileName, table, autoResizeColumns, culture);
        }

        /// <summary>
        /// Converts the list to excel table object.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">The list.</param>
        /// <returns></returns>
        public static ExcelTable ConvertListToExcelTableObject<T>(List<T> list)
        {
            if (list == null) throw new ArgumentNullException(nameof(list));

            Type type = typeof (T);
            DisplayNameAttribute displayNameAttribute = ReflectionUtils.GetCustomAttribute<DisplayNameAttribute>(type);
            string tableName = displayNameAttribute == null ? type.Name : displayNameAttribute.DisplayName;

            ExcelTable table = new ExcelTable {Name = tableName};

            //Add Column Names
            PropertyInfo[] properties = type.GetProperties();
            for (int i = 0; i < properties.Length; i++)
            {
                PropertyInfo property = properties[i];
                DisplayAttribute displayAttribute = ReflectionUtils.GetCustomAttribute<DisplayAttribute>(property);
                table.Columns.Add(displayAttribute == null ? property.Name : displayAttribute.Name);
            }

            //Add Row Values
            for (int i = 0; i < list.Count; i++)
            {
                T value = list[i];
                ExcelRow row = new ExcelRow();
                for (int j = 0; j < properties.Length; j++)
                {
                    PropertyInfo property = properties[j];
                    row[property.Name] = ReflectionHelper.GetPropertyValue(value, property.Name);
                }
                table.Rows.Add(row);
            }
            return table;
        }

        /// <summary>
        /// Writes the specified table to the excel file.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="table">The table.</param>
        /// <param name="autoResizeColumns">if set to <c>true</c> [automatic resize columns].</param>
        /// <param name="culture">The culture.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling")]
        public override void Write(string fileName, ExcelTable table, bool autoResizeColumns = false, CultureInfo culture = null)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentNullException(nameof(fileName));
            if (table == null) throw new ArgumentNullException(nameof(table));

            Application xlApp = null;
            Workbook xlWorkBook = null;
            Worksheet xlWorkSheet = null;
            object misValue = Missing.Value;

            CultureInfo oldCuture = null;

            try
            {
                if (culture != null)
                {
                    oldCuture = Thread.CurrentThread.CurrentCulture;
                    Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(culture.Name);
                }

                xlApp = new Application {Visible = false, ScreenUpdating = false};

                if (xlApp == null)
                {
                    throw new CoreLevelException("EXCEL could not be started. Check that your office installation and project references are correct.");
                }

                xlWorkBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                xlWorkSheet = xlWorkBook.Worksheets.Add();

                if (xlWorkSheet == null)
                {
                    throw new CoreLevelException("EXCEL could not be started. Check that your office installation and project references are correct.");
                }

                if (!string.IsNullOrEmpty(table.Name))
                {
                    xlWorkSheet.Name = table.Name;
                }
                List<ExcelRow> rows = table.Rows;
                List<string> columns = table.Columns;

                // add columns 
                for (int i = 0; i < columns.Count; i++)
                {
                    Range range = xlWorkSheet.Range["A1"].Offset[0, i];
                    range.Font.Bold = true;
                    //range.Interior.Color = ColorTranslator.ToOle(row.Color);
                    range.Value = columns[i];
                }

                // add data rows 
                for (int i = 0; i < rows.Count; i++)
                {
                    ExcelRow row = rows[i];
                    object[] values = row.Values.ToArray();
                    Range range = xlWorkSheet.Range["A2"].Offset[i].Resize[1, values.Length];
                    range.Font.Bold = row.Bold;
                    range.Font.Color = ColorTranslator.ToOle(row.Color);
                    range.Interior.Color = ColorTranslator.ToOle(row.BackColor);
                    range.NumberFormat = "@";
                    range.Value = values;
                }

                if (autoResizeColumns)
                {
                    xlWorkSheet.Columns.AutoFit();                    
                }

                xlWorkBook.SaveAs(fileName);
            }
            finally
            {
                try
                {
                    if (oldCuture != null)
                    {
                        Thread.CurrentThread.CurrentCulture = oldCuture;
                    }

                    if (xlWorkBook != null)
                    {
                        xlWorkBook.Close(misValue, misValue, misValue);
                    }
                    if (xlApp != null)
                    {
                        xlApp.Quit();
                    }
                }
                finally 
                {
                    ReleaseObject(xlWorkSheet);
                    ReleaseObject(xlWorkBook);
                    ReleaseObject(xlApp);
                }
            }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2001:AvoidCallingProblematicMethods", MessageId = "System.GC.Collect")]
        private static void ReleaseObject(object obj)
        {
            try
            {
                if(obj == null)
                {
                    return;
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;

                throw;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
