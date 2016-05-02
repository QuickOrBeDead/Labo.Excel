namespace Labo.Excel
{
    /// <summary>
    /// The excel read start operation info class.
    /// </summary>
    public sealed class ExcelReadStartOperationInfo
    {
        /// <summary>
        /// Gets the row count.
        /// </summary>
        /// <value>
        /// The row count.
        /// </value>
        public int RowCount { get; private set; }

        /// <summary>
        /// Gets the columns.
        /// </summary>
        /// <value>
        /// The columns.
        /// </value>
        public string[] Columns { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelReadStartOperationInfo"/> class.
        /// </summary>
        /// <param name="rowCount">The row count.</param>
        /// <param name="columns">The columns.</param>
        public ExcelReadStartOperationInfo(int rowCount, string[] columns)
        {
            RowCount = rowCount;
            Columns = columns;
        }
    }
}