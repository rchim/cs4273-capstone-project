using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace ExcelWrapper {
    /// <summary>
    /// Represents the interior of a single cell.
    /// </summary>
    public class Interior {
        private SLDocument _document;
        private Worksheet _worksheet;
        private int _row;
        private int _col;

        public Interior(int row, int col, Worksheet worksheet, SLDocument document) {
            this._document = document;
            this._worksheet = worksheet;
            this._row = row;
            this._col = col;
        }

        /// <summary>
        /// The fill color for the cell
        /// </summary>
        public int ColorIndex {
            set {
                // Setting a cell color on a worksheet activates that worksheet, 
                // if it wasn't already.
                _worksheet.Activate();

                SLStyle style = _document.GetCellStyle(_row, _col);
                style.Fill.SetPatternType(PatternValues.Solid);
                style.Fill.SetPatternForegroundColor(ColorUtils.excelColorIndexToColor(value));
                _document.SetCellStyle(_row, _col, style);
            }
        }

    }
}