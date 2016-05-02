using System.Collections.Generic;

namespace Labo.Excel
{
    /// <summary>
    /// The excel table class.
    /// </summary>
    public sealed class ExcelTable
    {
        private List<string> m_Columns;
        /// <summary>
        /// Gets or sets the columns.
        /// </summary>
        /// <value>
        /// The columns.
        /// </value>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public List<string> Columns
        {
            get { return m_Columns ?? (m_Columns = new List<string>()); }
            set { m_Columns = value; }
        }

        private List<ExcelRow> m_Rows;
        /// <summary>
        /// Gets or sets the rows.
        /// </summary>
        /// <value>
        /// The rows.
        /// </value>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public List<ExcelRow> Rows
        {
            get { return m_Rows ?? (m_Rows = new List<ExcelRow>()); }
            set { m_Rows = value; }
        }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; set; }
    }
}