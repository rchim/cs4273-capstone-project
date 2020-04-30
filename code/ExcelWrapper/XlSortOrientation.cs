using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWrapper {
    /// <summary>
    /// Specifies the sort orientation.
    /// </summary>
    public enum XlSortOrientation {
        /// <summary>
        /// Sorts by column
        /// </summary>
        xlSortColumns = 1,

        /// <summary>
        /// Sorts by row. This is the default value.
        /// </summary>
        xlSortRows = 2
    }
}
