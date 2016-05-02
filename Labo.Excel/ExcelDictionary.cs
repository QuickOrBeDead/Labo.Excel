using System;
using Labo.Common.Dynamic;

namespace Labo.Excel
{
    /// <summary>
    /// The excel dictionary class.
    /// </summary>
    public sealed class ExcelDictionary : DynamicDictionary
    {
        private ExcelDictionaryColumns m_Columns;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelDictionary"/> class.
        /// </summary>
        /// <param name="columns">The columns.</param>
        /// <exception cref="System.ArgumentNullException">columns</exception>
        public ExcelDictionary(ExcelDictionaryColumns columns)
        {
            if (columns == null)
            {
                throw new ArgumentNullException(nameof(columns));
            }

            m_Columns = columns;
        }

        /// <summary>
        /// Gets or sets the <see cref="System.Object"/> at the specified index.
        /// </summary>
        /// <value>
        /// The <see cref="System.Object"/>.
        /// </value>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public object this[int index]
        {
            get
            {
                string columnName = Columns[index];
                if (ContainsKey(columnName))
                {
                    return this[columnName];
                }
                return null;
            }
            set
            {
                string columnName = Columns[index];
                this[columnName] = value;
            }
        }

        /// <summary>
        /// Gets the columns.
        /// </summary>
        /// <value>
        /// The columns.
        /// </value>
        public ExcelDictionaryColumns Columns
        {
            get { return m_Columns ?? (m_Columns = new ExcelDictionaryColumns()); }
        }
    }
}