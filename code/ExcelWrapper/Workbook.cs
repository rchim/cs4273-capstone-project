using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;

namespace ExcelWrapper {
    public class Workbook {
        private const int NUM_DEFAULT_SHEETS = 3;
        
        private Application _application;
        private SLDocument _document;
        private Sheets _sheets;
        private bool _closed = false;

        /// <summary>
        /// Creates or opens a workbook.
        /// </summary>
        /// <param name="application">
        /// The Excel application managing this workbook.
        /// </param>
        /// <param name="filePath">
        /// A file path to open an existing workbook, or null
        /// to create a new one.
        /// </param>
        public Workbook(Application application, string filePath = null) {
            this._application = application;
            
            if (filePath != null) {
                this._document = new SLDocument(filePath);
            }
            else {
                this._document = new SLDocument();

                // Since we are making a new workbook, we start with
                // some default sheets, which are named Sheet1, Sheet2, etc.
                // Sheet1 can come from Spreadsheet Light's default.
                this._document.RenameWorksheet(
                    SLDocument.DefaultFirstSheetName,
                    "Sheet1"
                );
                for (int i = 2; i <= NUM_DEFAULT_SHEETS; i++) {
                    this._document.AddWorksheet("Sheet" + i);
                }
            }

            // Finally, build the Sheets collection from the document we 
            // created/opened.
            this._sheets = new Sheets(this);
        }

        /// <summary>
        /// Saves the workbook to a file. Currently only file format 56 (for xlsx)
        /// is implemented. Attempting to save with a different file format will 
        /// result in an error message being written to console, with no save performed.
        /// </summary>
        /// <param name="name">
        /// The file path and name.
        /// </param>
        /// <param name="fileFormat">
        /// 56 for xlsx; no other formats are implemented yet.
        /// </param>
        public void SaveAs(string fileName, int fileFormat) {
            if (fileFormat != 56) {
                Console.Error.WriteLine("Tried to save workbook to " +
                    $"unimplemented file format {fileFormat}. " +
                    "Save not performed.");
                return;
            }

            _document.SaveAs(fileName);

            // Saving to a file with SpreadsheetLight clears the document object. 
            // But we aren't necessarily done, so we reopen the document.
            _document = new SLDocument(fileName);
        }

        /// <summary>
        /// Closes this workbook, disposing of its resources.
        /// </summary>
        public void Close() {
            if (!_closed) {
                _document.CloseWithoutSaving();
                _closed = true;
            }            
        }

        public void Activate() {
            this._application.ActiveWorkbook = this;
        }

        /// <summary>
        /// Returns all the worksheets in the workbook.
        /// </summary>
        public Sheets Worksheets() {
            return _sheets;
        }

        /// <summary>
        /// Returns a single worksheet in the workbook, by index.
        /// </summary>
        public Worksheet Worksheets(int index) {
            return this.Worksheets().Item(index);
        }

        /// <summary>
        /// Returns a single worksheet in the workbook, by name.
        /// </summary>
        public Worksheet Worksheets(string name) {
            return this.Worksheets().Item(name);
        }
        
        /// <summary>
        /// Returns the currently selected worksheet in this workbook.
        /// </summary>
        public Worksheet ActiveSheet {
            get { 
                return new Worksheet(
                    _document.GetCurrentWorksheetName(), 
                    this
                ); 
            }
        }

        /// <summary>
        /// The SpreadsheetLight document object that represents this
        /// workbook behind the scenes.
        /// </summary>
        public SLDocument Document {
            get { return _document; }
        }
    }
}
