using System;
using System.Collections.Generic;
using System.Linq;

namespace Labo.Excel
{
    /// <summary>
    /// The extensions class.
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// To the excel table.
        /// </summary>
        /// <param name="excelDictionaryCollection">The excel dictionary collection.</param>
        /// <returns></returns>
        public static ExcelTable ToExcelTable(this ExcelDictionaryCollection excelDictionaryCollection)
        {
            if (excelDictionaryCollection == null) throw new ArgumentNullException(nameof(excelDictionaryCollection));

            ExcelTable table = new ExcelTable
                                   {
                                       Columns = excelDictionaryCollection.Columns.ToList()
                                   };
            for (int i = 0; i < excelDictionaryCollection.Count; i++)
            {
                var row = excelDictionaryCollection[i];
                ExcelRow excelRow = new ExcelRow();
                ExcelDictionaryColumns excelDictionaryColumns = row.Columns;
                for (int j = 0; j < excelDictionaryColumns.Count; j++)
                {
                    string columnName = excelDictionaryColumns[j];
                    excelRow[columnName] = row[columnName];
                }

                table.Rows.Add(excelRow);
            }

            return table;
        }

        /// <summary>
        /// To the excel table.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">The list.</param>
        /// <returns></returns>
        public static ExcelTable ToExcelTable<T>(List<T> list)
        {
            return OfficeInteropExcelReaderWriter.ConvertListToExcelTableObject(list);
        }
    }
}
