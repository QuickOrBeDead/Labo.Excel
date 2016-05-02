using System;
using System.Collections.Generic;
using System.Globalization;

namespace Labo.Excel
{
    /// <summary>
    /// The excel reader writer interface.
    /// </summary>
    public interface IExcelReaderWriter
    {
        /// <summary>
        /// Reads the specified excel.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns></returns>
        ExcelDictionaryCollection Read(string fileName);

        /// <summary>
        /// Reads the specified excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="mapper">The mapper.</param>
        /// <param name="onStartReading">The on start reading.</param>
        /// <returns></returns>
        List<T> Read<T>(string fileName, Func<ExcelRowReadOperation, T> mapper, Action<ExcelReadStartOperationInfo> onStartReading = null);

        /// <summary>
        /// Writes the specified table to the excel file.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="table">The table.</param>
        /// <param name="autoResizeColumns">if set to <c>true</c> [automatic resize columns].</param>
        /// <param name="culture">The culture.</param>
        void Write(string fileName, ExcelTable table, bool autoResizeColumns = false, CultureInfo culture = null);

        /// <summary>
        /// Writes the specified table to the excel file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="list">The list.</param>
        /// <param name="autoResizeColumns">if set to <c>true</c> [automatic resize columns].</param>
        /// <param name="culture">The culture.</param>
        void Write<T>(string fileName, List<T> list, bool autoResizeColumns = false, CultureInfo culture = null);

        /// <summary>
        /// Writes the specified table to the excel file.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="excelDictionaryCollection">The excel dictionary collection.</param>
        /// <param name="autoResizeColumns">if set to <c>true</c> [automatic resize columns].</param>
        /// <param name="culture">The culture.</param>
        void Write(string fileName, ExcelDictionaryCollection excelDictionaryCollection, bool autoResizeColumns = false, CultureInfo culture = null);
    }
}