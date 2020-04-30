using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWrapper {
    /// <summary>
    /// A collection of the currently open workbooks. This library
    /// only allows one open workbook at a time for simplicity. The
    /// one open workbook is the application's active workbook.
    /// </summary>
    public class Workbooks {
        private readonly Application _application;
        private Workbook _openWorkbook;

        public Workbooks(Application application) {
            this._application = application;
            this._openWorkbook = null;
        }

        /// <summary>
        /// Create a new workbook and activate it. This closes the previously
        /// open workbook.
        /// </summary>
        public Workbook Add() {
            this._openWorkbook = new Workbook(this._application);
            this._openWorkbook.Activate();
            return this._openWorkbook;
        }

        /// <summary>
        /// Open a workbook from file and activate it. This closes the previously
        /// open workbook.
        /// </summary>
        public Workbook Open(string path) {
            this._openWorkbook = new Workbook(this._application, path);
            this._application.ActiveWorkbook = this._openWorkbook;
            return this._openWorkbook;
        }
         
        public Application Application {
            get { return this._application; }
        }

        /// <summary>
        /// Disposes of all the workbooks.
        /// </summary>
        internal void CloseAll() {
            // Since we are only allowing one workbook, we just close the open
            // workbook.
            if (_openWorkbook != null) {
                _openWorkbook.Close();
                _openWorkbook = null;
            }
        }
    }
}
