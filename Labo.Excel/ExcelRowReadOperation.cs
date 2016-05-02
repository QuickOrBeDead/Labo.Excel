namespace Labo.Excel
{
    /// <summary>
    /// The excel read opertation class.
    /// </summary>
    public sealed class ExcelRowReadOperation
    {
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="ExcelRowReadOperation"/> is cancel.
        /// </summary>
        /// <value>
        ///   <c>true</c> if cancel; otherwise, <c>false</c>.
        /// </value>
        public bool Cancel { get; set; }

        /// <summary>
        /// Gets the row.
        /// </summary>
        /// <value>
        /// The row.
        /// </value>
        public ExcelDictionary Row { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelRowReadOperation"/> class.
        /// </summary>
        /// <param name="row">The row.</param>
        public ExcelRowReadOperation(ExcelDictionary row)
        {
            Row = row;
            Cancel = false;
        }
    }
}