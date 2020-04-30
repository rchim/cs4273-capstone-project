using SpreadsheetLight;
using System.Collections.Generic;
using System.Linq;

namespace ExcelWrapper {
    /// <summary>
    /// Represents all the sheets in a workbook.
    /// </summary>
    public class Sheets {
        private Workbook _workbook;

        /// <summary>
        /// Creates a Sheets object that represents the worksheets for a
        /// particular workbook.
        /// </summary>
        public Sheets(Workbook workbook) {
            this._workbook = workbook;
        }

        /// <summary>
        /// Adds a new worksheet AND activates it.
        /// </summary>
        public Worksheet Add() {
            List<string> existingNames = _workbook.Document.GetSheetNames();
            
            int number = 1;
            string tentativeName = "Sheet1";
            while (existingNames.Contains(tentativeName)) {
                number++;
                tentativeName = "Sheet" + number;
            }
            _workbook.Document.AddWorksheet(tentativeName);

            return new Worksheet(tentativeName, _workbook);
        }

        /// <summary>
        /// Returns the worksheet with the given name (or throws
        /// an exception if it doesn't exist).
        /// </summary>
        public Worksheet Item(string name) {
            if (!_workbook.Document.GetSheetNames().Contains(name)) {
                throw new System.Exception(
                    $"Could not find worksheet named '{name}'");
            }

            return new Worksheet(name, _workbook);
        }

        /// <summary>
        /// Returns the worksheet at the given index.
        /// </summary>
        public Worksheet Item(int index) {
            string name = _workbook.Document.GetSheetNames()[index-1];
            return this.Item(name);
        }

        /// <summary>
        /// The number of worksheets in this collection of worksheets.
        /// </summary>
        public int Count {
            get { return _workbook.Document.GetSheetNames().Count; }
        }
    }
}