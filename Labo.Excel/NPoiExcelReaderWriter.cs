using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using Labo.Common.Utils;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Labo.Excel
{
    /// <summary>
    /// The npoi excel reader writer class.
    /// </summary>
    public class NPoiExcelReaderWriter : BaseExcelReaderWriter
    { 
        /// <summary>
        /// Reads the specified excel internal method.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="mapper">The mapper.</param>
        /// <param name="onStartReading">The on start reading.</param>
        /// <returns></returns>
        public override List<T> ReadInternal<T>(string fileName, Func<ExcelRowReadOperation, T> mapper, Action<ExcelReadStartOperationInfo> onStartReading = null)
        {
            string extension = Path.GetExtension(fileName).ToUpperInvariant();
            using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook;
                if (extension == ".XLSX")
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                else
                {
                    workbook = new HSSFWorkbook(fileStream);
                }

                ISheet sheet = workbook.GetSheetAt(0);
                IRow headerRow = sheet.GetRow(0);

                int columnsLength = headerRow.LastCellNum;

                ExcelDictionaryColumns excelDictionaryColumns = new ExcelDictionaryColumns();
                string[] columns = new string[columnsLength];
                for (int i = headerRow.FirstCellNum; i < columns.Length; i++)
                {
                    string columnName = ConvertUtils.ChangeType<string>(headerRow.GetCell(i));
                    columns[i] = columnName;
                    excelDictionaryColumns.Add(columnName);
                }

                excelDictionaryColumns.EnsureColumnsNotNull();
                excelDictionaryColumns.EnsureColumnsAreUnique();

                int lastRowNum = sheet.LastRowNum;
                if (onStartReading != null)
                {
                    onStartReading(new ExcelReadStartOperationInfo(lastRowNum, columns));
                }

                return ReadInternal(mapper, lastRowNum, excelDictionaryColumns, sheet, columns);
            }
        }

        private static List<T> ReadInternal<T>(Func<ExcelRowReadOperation, T> mapper, int lastRowNum, ExcelDictionaryColumns excelDictionaryColumns, ISheet sheet, string[] columns)
        {
            SortedList<int, T> result = new SortedList<int, T>(Comparer<int>.Default);

            for(int i = 1; i <= lastRowNum; i++)
            {
                ExcelDictionary row = new ExcelDictionary(excelDictionaryColumns);
                IRow excelRow = sheet.GetRow(i);
                if (excelRow == null)
                {
                    break;
                }
                FillExcelRow(columns, excelRow, row);
                ExcelRowReadOperation readOperation = new ExcelRowReadOperation(row);
                T item = mapper(readOperation);
                if (readOperation.Cancel)
                {
                    break;
                }

                result[i] = item;
            }

            return result.Values.ToList();
        }

        private static void FillExcelRow(string[] columns, IRow excelRow, IDictionary<string, object> row)
        {
            for (int j = excelRow.FirstCellNum; j < columns.Length; j++)
            {
                if (excelRow.GetCell(j) != null)
                {
                    row[columns[j]] = ConvertUtils.ChangeType<string>(excelRow.GetCell(j));
                }
            }
        }

        private readonly ConcurrentDictionary<string, ICellStyle> m_RowStyleCache = new ConcurrentDictionary<string, ICellStyle>();

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
            Write(fileName, OfficeInteropExcelReaderWriter.ConvertListToExcelTableObject(list), autoResizeColumns, culture);
        }

        /// <summary>
        /// Writes the specified table to the excel file.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="table">The table.</param>
        /// <param name="autoResizeColumns">if set to <c>true</c> [automatic resize columns].</param>
        /// <param name="culture">The culture.</param>
        public override void Write(string fileName, ExcelTable table, bool autoResizeColumns = false, CultureInfo culture = null)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentNullException(nameof(fileName));
            if (table == null) throw new ArgumentNullException(nameof(table));

            if (culture == null)
            {
                culture = Thread.CurrentThread.CurrentCulture;
            }

            string extension = Path.GetExtension(fileName).ToUpperInvariant();
            IWorkbook workbook;
            if (extension == ".XLSX")
            {
                workbook = new XSSFWorkbook();
               
            }
            else
            {
                workbook = new HSSFWorkbook();                
            }
            ISheet sheet = workbook.CreateSheet(table.Name ?? "sheet1");

            List<ExcelRow> rows = table.Rows;
            List<string> columns = table.Columns;

            int rowIndex = 0;
            if (columns.Count > 0)
            {
                IRow headerRow = sheet.CreateRow(0);
                ICellStyle headerStyle = GetHeaderStyle(workbook);
                rowIndex++;

                // add columns 
                for (int i = 0; i < columns.Count; i++)
                {
                    ICell cell = headerRow.CreateCell(i);
                    cell.SetCellValue(columns[i]);

                    cell.CellStyle = headerStyle;
                }
            }

            if (rows.Count > 0)
            {
                // add data rows 
                for (int i = 0; i < rows.Count; i++, rowIndex++)
                {
                    IRow row = sheet.CreateRow(rowIndex);

                    ExcelRow excelRow = rows[i];

                    IList<object> values = excelRow.Values.ToList();
                    for (int j = 0; j < values.Count; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        object value = values[j];
                        if(value is byte || value is int || value is short || value is long || value is float || value is decimal || value is double)
                        {
                            cell.SetCellValue(ConvertUtils.ChangeType<double>(value, culture));
                            cell.SetCellType(CellType.Numeric);
                        }
                        else if (value is DateTime)
                        {
                            cell.SetCellValue((DateTime)value);
                        }
                        else if (value is bool)
                        {
                            cell.SetCellValue((bool)value);
                            cell.SetCellType(CellType.Boolean);
                        }
                        else if (value == null)
                        {
                            cell.SetCellValue((string)null);
                            cell.SetCellType(CellType.Blank);
                        }
                        else
                        {
                            cell.SetCellValue(Convert.ToString(value, culture));
                            cell.SetCellType(CellType.String);
                        }
                      
                        cell.CellStyle = this.GetRowStyle(workbook, excelRow);
                    }
                }
            }

            if (autoResizeColumns)
            {
                // auto size columns
                for (int i = 0; i < columns.Count; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
            }

            using (FileStream fileData = new FileStream(fileName, FileMode.Create))
            {
                workbook.Write(fileData);
            }
        }

        private ICellStyle GetRowStyle(IWorkbook workbook, ExcelRow excelRow)
        {
            short colorIndex;
            short? backGroundColorIndex = null;
            if(workbook is HSSFWorkbook)
            {
                colorIndex = GetColorIndex((HSSFWorkbook)workbook, excelRow.Color);
                if (!excelRow.BackColor.IsEmpty)
                {
                    backGroundColorIndex = GetColorIndex((HSSFWorkbook)workbook, excelRow.BackColor);
                }
            }
            else
            {
                colorIndex = GetColorIndex(new HSSFWorkbook(), excelRow.Color);
                if (!excelRow.BackColor.IsEmpty)
                {
                    backGroundColorIndex = GetColorIndex(new HSSFWorkbook(), excelRow.BackColor);
                }
            }

            bool bold = excelRow.Bold;

            return m_RowStyleCache.GetOrAdd(string.Format(CultureInfo.InvariantCulture, "{0}-{1}-{2}", colorIndex, backGroundColorIndex, bold), x => CreateCellStyle(workbook, colorIndex, backGroundColorIndex, bold));
        }

        private static short GetColorIndex(HSSFWorkbook workbook, Color color)
        {
            HSSFPalette palette = workbook.GetCustomPalette();
            return palette.FindColor(color.R, color.G, color.B).Indexed;
        }

        private static ICellStyle CreateCellStyle(IWorkbook workbook, short colorIndex, short? backGroundColorIndex, bool bold)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();

            if (backGroundColorIndex.HasValue)
            {
                cellStyle.FillForegroundColor = backGroundColorIndex.Value;
                cellStyle.FillPattern = FillPattern.SolidForeground;
                cellStyle.FillBackgroundColor = backGroundColorIndex.Value;
            }
        
            short formatId = HSSFDataFormat.GetBuiltinFormat("text");
            if (formatId == -1)
            {
                IDataFormat newDataFormat = workbook.CreateDataFormat();
                cellStyle.DataFormat = newDataFormat.GetFormat("text");
            }
            else
            {
                cellStyle.DataFormat = formatId;
            }

            IFont rowFont = workbook.CreateFont();
            if (bold)
            {
                rowFont.Boldweight = (short)FontBoldWeight.Bold;
            }
            rowFont.Color = colorIndex;
            cellStyle.SetFont(rowFont);
            return cellStyle;
        }

        private static ICellStyle GetHeaderStyle(IWorkbook workbook)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            IFont cellFont = workbook.CreateFont();
            cellFont.Boldweight = (short)FontBoldWeight.Bold;
            cellStyle.SetFont(cellFont);
            return cellStyle;
        }
    }
}
