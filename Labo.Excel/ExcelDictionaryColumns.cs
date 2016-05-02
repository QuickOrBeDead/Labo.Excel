using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Labo.Common.Exceptions;

namespace Labo.Excel
{
    /// <summary>
    /// The excel dictionary columns class.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public sealed class ExcelDictionaryColumns : List<string>
    {
        /// <summary>
        /// Ensures the columns not null.
        /// </summary>
        /// <exception cref="UserLevelException"></exception>
        public void EnsureColumnsNotNull()
        {
            List<string> nullColumnNos = new List<string>();
            for (int i = 0; i < Count; i++)
            {
                string column = this[i];
                if(string.IsNullOrWhiteSpace(column))
                {
                    nullColumnNos.Add((i+1).ToString(CultureInfo.InvariantCulture));
                }
            }
            if (nullColumnNos.Count > 0)
            {
                throw new UserLevelException(string.Format(CultureInfo.CurrentCulture, "Excel Kolon Baþlýklarý Boþ Olamaz. Kolonlar Numaralarý: {0}", string.Join(",", nullColumnNos)));
            }
        }

        /// <summary>
        /// Ensures the columns are unique.
        /// </summary>
        /// <exception cref="UserLevelException"></exception>
        public void EnsureColumnsAreUnique()
        {
            string[] dublicateColumns = (from column in this
                                         group column by column
                                         into columnGroup
                                         where columnGroup.Count() > 1
                                         select columnGroup.Key).ToArray();
            if (dublicateColumns.Length > 0)
            {
                throw new UserLevelException(string.Format(CultureInfo.CurrentCulture, "Ayný Ýsimli Excel Kolon Baþlýklarý Tanýmlanamaz. Kolonlar: {0}", string.Join(",", dublicateColumns.Select(x => "'{0}'"))));
            }
        }
    }
}