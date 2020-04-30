using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;

namespace ExcelWrapper {
    /// <summary>
    /// Represents an Excel worksheet.
    /// </summary>
    public class Worksheet {
        private string _name;
        private Workbook _workbook;

        /// <summary>
        /// The row number where the worksheet is split into panes (the number of 
        /// rows above the split).
        /// </summary>
        private int _splitRow = 0;

        /// <summary>
        /// The column number where the worksheet is split into panes (the number of
        /// columns above the split).
        /// </summary>
        private int _splitColumn = 0;

        public Worksheet(string name, Workbook workbook) {
            this._name = name;
            this._workbook = workbook;
        }

        /// <summary>
        /// Makes this worksheet the active worksheet. This is equivalent to
        /// choosing the sheet's tab.
        /// </summary>
        public void Activate() {
            this._workbook.Document.SelectWorksheet(_name);
        }

        /// <summary>
        /// Deletes the worksheet. Note that the active worksheet cannot be
        /// deleted.
        /// </summary>
        public void Delete() {
            // In the future, may want to check that the sheet
            // with this name hasn't been renamed or deleted by 
            // another Worksheet instance. More importantly, probably want to
            // enforce one worksheet instance per name-workbook combination.
            // In other words, one worksheet instance per actual worksheet.
            _workbook.Document.DeleteWorksheet(_name);
        }

        /// <summary>
        /// Returns a Range object that represents a cell or range of cells.
        /// </summary>
        /// <param name="rangeString">
        /// For example, "A2:IV65536" or "B3"
        /// </param>
        public Range Range(string rangeString) {
            return new Range(rangeString, this, _workbook.Document);
        }

        /// <summary>
        /// Returns a Range object that represents all the columns 
        /// (entire columns) in the worksheet.
        /// </summary>
        public Range Columns() {
            return new Range(RangeType.Columns, this, _workbook.Document);
        }

        /// <summary>
        /// Returns a single column in the worksheet.
        /// </summary>
        /// <param name="index">
        /// The first column is 1, the second column is 2...
        /// </param>
        public Range Columns(int index) {
            return this.Columns().Item(index);
        }

        /// <summary>
        /// Returns a range made up of columns. For example,
        /// Columns("AB") or Columns("A:IV").
        /// </summary>
        public Range Columns(string columnNameRange) {
            return ExcelWrapper.Range.FromColumnRange(
                columnNameRange,
                this,
                _workbook.Document
            );
        }

        /// <summary>
        /// Returns a range object that represents all the rows (entire rows)
        /// in the worksheet.
        /// </summary>
        public Range Rows() {
            return new Range(RangeType.Rows, this, _workbook.Document);
        }

        /// <summary>
        /// Returns a single row in the worksheet.
        /// </summary>
        /// <param name="index">
        /// The first row is 1, the second row is 2...
        /// </param>
        public Range Rows(int index) {
            return this.Rows().Item(index);
        }

        /// <summary>
        /// Returns a range consisting of all the cells in the sheet. 
        /// </summary>
        public Range Cells() {
            return new Range(RangeType.Cells, this, _workbook.Document);
        }

        /// <summary>
        /// Returns a single cell in the sheet.
        /// </summary>
        public Range Cells(int row, int col) {
            return this.Cells().Item(row, col);
        }

        /// <summary>
        /// Returns the number of rows from the first nonempty row to
        /// the last nonempty row (inclusive). This matches the behavior
        /// of wksht.UsedRange.Rows.Count in the real Excel library. We 
        /// implement it like this to avoid having to match Excel's 
        /// UsedRange functionality.
        /// </summary>
        public int UsedRangeRowsCount() {
            string prevSelectedWorksheet = _workbook.Document.GetCurrentWorksheetName();
            _workbook.Document.SelectWorksheet(_name);
            int result = _workbook.Document.GetWorksheetStatistics().NumberOfRows;
            _workbook.Document.SelectWorksheet(prevSelectedWorksheet);
            return result;
        }

        /// <summary>
        /// Protects the worksheet with a password. Currently, this method is
        /// not implemented (SpreadsheetLight doesn't provide this capability).
        /// Attempting to protect a worksheet results in an error message being
        /// written to console and no change to the worksheet. To anyone
        /// implementing this method in the future, consider saving the
        /// workbook to xlsx and then applying protection by other means.
        /// If necessary, the same technique can be used to remove the password at a
        /// later time, at which point SpreadsheetLight will be able to access the
        /// worksheet again.
        /// </summary>
        public void Protect(string password) {
            Console.Error.WriteLine("Tried to protect worksheet with password," +
                " but this functionality isn't implemented yet." +
                " Protection not performed.");
        }

        /// <summary>
        /// The worksheet name as a string.
        /// </summary>
        public string Name {
            get {
                return _name;
            }
            set {
                _workbook.Document.RenameWorksheet(this._name, value);
                _name = value;
            }
        }

        /// <summary>
        /// The row number where the worksheet is split into panes (the number of 
        /// rows above the split).
        /// </summary>
        public int SplitRow {
            set {
                _splitRow = value;

                string prevSelectedWorksheet = _workbook.Document.GetCurrentWorksheetName();
                _workbook.Document.SelectWorksheet(_name);
                _workbook.Document.SplitPanes(_splitRow, _splitColumn, true);
                _workbook.Document.SelectWorksheet(prevSelectedWorksheet);
            }
        }

        /// <summary>
        /// The column number where the worksheet is split into panes (the number of 
        /// columns above the split).
        /// </summary>
        public int SplitColumn {
            set {
                _splitColumn = value;

                string prevSelectedWorksheet = _workbook.Document.GetCurrentWorksheetName();
                _workbook.Document.SelectWorksheet(_name);
                _workbook.Document.SplitPanes(_splitRow, _splitColumn, true);
                _workbook.Document.SelectWorksheet(prevSelectedWorksheet);
            }
        }

        /// <summary>
        /// True if split panes are frozen.
        /// </summary>
        public bool FreezePanes {
            set {
                string prevSelectedWorksheet = _workbook.Document.GetCurrentWorksheetName();
                _workbook.Document.SelectWorksheet(_name);

                if (value) {
                    _workbook.Document.FreezePanes(_splitRow, _splitColumn);
                } else {
                    _workbook.Document.UnfreezePanes();

                    // In SpreadsheetLight, unfreezing panes also gets rid of the
                    // split. We *don't* want this behavior, so we restore the split.
                    _workbook.Document.SplitPanes(_splitRow, _splitColumn, true);
                }

                _workbook.Document.SelectWorksheet(prevSelectedWorksheet);
            }
        }
    }
}
