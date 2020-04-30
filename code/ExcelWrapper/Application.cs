using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWrapper
{
    /// <summary>
    /// Represents the entire Microsoft Excel application.
    /// </summary>
    public class Application
    {
        private Workbooks _workbooks;
        private Workbook _activeWorkbook;        

        public Application() {
            this._workbooks = new Workbooks(this);
            this._activeWorkbook = null;
        }

        /// <summary>
        /// Returns the worksheets in the active workbook.
        /// </summary>
        public Sheets Worksheets() {
            return _activeWorkbook.Worksheets();
        }

        /// <summary>
        /// Returns a single worksheet in the active workbook, by index.
        /// </summary>
        public Worksheet Worksheets(int index) {
            return _activeWorkbook.Worksheets(index);
        }

        /// <summary>
        /// Returns a single worksheet in the active workbook, by name.
        /// </summary>
        public Worksheet Worksheets(string name) {
            return _activeWorkbook.Worksheets(name);
        }

        /// <summary>
        /// Closes all the workbooks being used by the application,
        /// and disposes of their resources.
        /// </summary>
        public void Quit() {
            this._activeWorkbook = null;
            _workbooks.CloseAll();
        }

        public Workbooks Workbooks {
            get { return _workbooks; }
        }

        public Workbook ActiveWorkbook {
            get { return _activeWorkbook; }
            set { _activeWorkbook = value; }
        }

        public Worksheet ActiveWorksheet {
            get { return this.ActiveWorkbook.ActiveSheet; }
        }
    }    
}
