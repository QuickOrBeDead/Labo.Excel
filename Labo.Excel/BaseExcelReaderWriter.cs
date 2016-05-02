using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Labo.Common.Exceptions;
using Labo.Excel;

namespace Labo.Excel
{
    /// <summary>
    /// The base excel reader writer class.
    /// </summary>
    public abstract class BaseExcelReaderWriter : IExcelReaderWriter
    {
        /// <summary>
        /// Reads the specified excel.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns></returns>
        public ExcelDictionaryCollection Read(string fileName)
        {
            string[] columns = null;
            List<ExcelDictionary> excelDictionaryList = Read(fileName, x => x.Row, x => columns = x.Columns);
            if (columns == null)
            {
                columns = new string[0];
            }

            return new ExcelDictionaryCollection(excelDictionaryList, columns);
        }

        /// <summary>
        /// Reads the specified excel internal method.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="mapper">The mapper.</param>
        /// <param name="onStartReading">The on start reading.</param>
        /// <returns></returns>
        public abstract List<T> ReadInternal<T>(string fileName, Func<ExcelRowReadOperation, T> mapper, Action<ExcelReadStartOperationInfo> onStartReading = null);

        /// <summary>
        /// Writes the specified table to the excel file.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="table">The table.</param>
        /// <param name="autoResizeColumns">if set to <c>true</c> [automatic resize columns].</param>
        /// <param name="culture">The culture.</param>
        public abstract void Write(string fileName, ExcelTable table, bool autoResizeColumns = false, CultureInfo culture = null);

        /// <summary>
        /// Writes the specified table to the excel file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="list">The list.</param>
        /// <param name="autoResizeColumns">if set to <c>true</c> [automatic resize columns].</param>
        /// <param name="culture">The culture.</param>
        public abstract void Write<T>(string fileName, List<T> list, bool autoResizeColumns = false, CultureInfo culture = null);

        /// <summary>
        /// Writes the specified table to the excel file.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="excelDictionaryCollection">The excel dictionary collection.</param>
        /// <param name="autoResizeColumns">if set to <c>true</c> [automatic resize columns].</param>
        /// <param name="culture">The culture.</param>
        public void Write(string fileName, ExcelDictionaryCollection excelDictionaryCollection, bool autoResizeColumns = false, CultureInfo culture = null)
        {
            Write(fileName, excelDictionaryCollection.ToExcelTable(), autoResizeColumns, culture);
        }

        /// <summary>
        /// Reads the specified excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="mapper">The mapper.</param>
        /// <param name="onStartReading">The on start reading.</param>
        /// <returns></returns>
        public List<T> Read<T>(string fileName, Func<ExcelRowReadOperation, T> mapper, Action<ExcelReadStartOperationInfo> onStartReading = null)
        {
            try
            {
                return ReadInternal(fileName, mapper, onStartReading);
            }
            catch (IOException ioException)
            {
                throw new CriticalUserLevelException(string.Format(CultureInfo.CurrentCulture, "Dosya Erişim Hatası: '{0}'", fileName), ioException);
            }
        }
    }
}
