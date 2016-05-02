using System.Drawing;
using Labo.Common.Dynamic;

namespace Labo.Excel
{
    /// <summary>
    /// The excel row class.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public sealed class ExcelRow : DynamicDictionary
    {
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="ExcelRow"/> is bold.
        /// </summary>
        /// <value>
        ///   <c>true</c> if bold; otherwise, <c>false</c>.
        /// </value>
        public bool Bold { get; set; }

        /// <summary>
        /// Gets or sets the color.
        /// </summary>
        /// <value>
        /// The color.
        /// </value>
        public Color Color { get; set; }

        /// <summary>
        /// Gets or sets the color of the back.
        /// </summary>
        /// <value>
        /// The color of the back.
        /// </value>
        public Color BackColor { get; set; }
    }
}