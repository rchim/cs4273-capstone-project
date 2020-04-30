using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;
using ADODB;

namespace ExcelWrapper {
    /// <summary>
    /// Represents a cell, a row, a column, or a contiguous
    /// block of cells.
    /// </summary>
    public class Range {
        /// <summary>
        /// The value for this is based on the code in ErrFunctions.vb and
        /// Import_xls.aspx.vb.
        /// </summary>
        private static readonly int ROW_INDEX_INFINITY = 65536;

        /// <summary>
        /// The value for this is based on the code in ErrFunctions.vb and
        /// Import_xls.aspx.vb.
        /// </summary>
        private static readonly int COL_INDEX_INFINITY = 
            CoordinateUtils.ToColumnIndex("IV");

        /// <summary>
        /// It's important that a date get a date format or else it will appear in the
        /// spreadsheet as a number. Also, when reading cell values programmatically,
        /// we rely on the format code to distinguish between dates and numbers.
        /// This particular format was chosen because it appears in Import_xls.aspx.vb,
        /// and is the only date format code to appear in either Import_xls or 
        /// ErrFunctions.vb. Ideally, if we don't want to mistake dates for numbers 
        /// when reading them, this should be the only format code we ever apply to dates.
        /// </summary>
        private static readonly string DATE_FORMAT_CODE = "dd/mm/yyyy"; 

        private SLDocument _document;
        private Worksheet _worksheet;
        private RangeType _type;
        private int _minRowIndex;
        private int _minColIndex;
        private int _maxRowIndex;        
        private int _maxColIndex;

        /// <summary>
        /// Initializes a range.
        /// </summary>
        public Range(
            int minRowindex,
            int minColIndex,
            int maxRowIndex,
            int maxColIndex,
            RangeType type,
            Worksheet worksheet,
            SLDocument document
        ) {
            this._minRowIndex = minRowindex;
            this._minColIndex = minColIndex;
            this._maxRowIndex = maxRowIndex;
            this._maxColIndex = maxColIndex;
            this._type = type;
            this._worksheet = worksheet;
            this._document = document;
        }

        /// <summary>
        /// Initializes a range consisting of a single cell, viewed as a cell.
        /// </summary>
        public Range(
            int rowIndex,
            int colIndex,
            Worksheet worksheet,
            SLDocument document
        ) : this(
            rowIndex,
            colIndex,
            rowIndex,
            colIndex,
            RangeType.Cells,
            worksheet,
            document
        ) { }

        /// <summary>
        /// Initializes a Range made up of cells, based on a range string
        /// like "A1" or "B2:LC65536".
        /// </summary>
        public Range(
            string rangeString, 
            Worksheet worksheet, 
            SLDocument document
        ) {
            this._document = document;
            this._worksheet = worksheet;
            this._type = RangeType.Cells;

            (_minRowIndex, _minColIndex, _maxRowIndex, _maxColIndex) =
                    CoordinateUtils.RangeStringToCoords(rangeString);
        }

        /// <summary>
        /// Initializes a Range that spans an entire worksheet.
        /// </summary>
        public Range(
            RangeType type,
            Worksheet worksheet, 
            SLDocument document
        ) {
            this._document = document;
            this._worksheet = worksheet;
            this._type = type;

            // The following indices mean that this range spans the
            // entire sheet.
            _minRowIndex = 1;
            _minColIndex = 1;
            _maxRowIndex = ROW_INDEX_INFINITY;
            _maxColIndex = COL_INDEX_INFINITY;
        }

        /// <summary>
        /// Returns a range made of columns. For example, 
        /// FromColumnRange("AB") or FromColumnRange("B:IV").
        /// </summary>
        public static Range FromColumnRange(
            string columnNameRange,
            Worksheet worksheet,
            SLDocument document
        ) {
            string[] colNames = columnNameRange.Contains(":") ?
                    columnNameRange.Split(new char[] { ':' }) :
                    new string[] { columnNameRange, columnNameRange };

            int minColIndex = CoordinateUtils.ToColumnIndex(colNames[0]);
            int maxColIndex = CoordinateUtils.ToColumnIndex(colNames[1]);

            return new ExcelWrapper.Range(
                1, 
                minColIndex, 
                ROW_INDEX_INFINITY, 
                maxColIndex,
                RangeType.Columns,
                worksheet,
                document
            );
        }

        /// <summary>
        /// Returns the cell at the specified offset from the top
        /// left cell of the range.
        /// </summary>
        /// <param name="rowOffset">
        /// The 1-based offset from the first row of the range in the
        /// direction top-to-bottom.
        /// </param>
        /// <param name="colOffset">
        /// The 1-based offset from the first column of the range in
        /// the direction left-to-right.
        /// </param>
        public Range Item(int rowOffset, int colOffset) {
            if (rowOffset < 1 || colOffset < 1) {
                throw new ArgumentOutOfRangeException("rowOffset and colOffset " +
                    $"must both be positive, but received {rowOffset}, {colOffset}.");
            }

            int newRowIndex = _minRowIndex + rowOffset - 1;
            int newColIndex = _minColIndex + colOffset - 1;

            return new Range(newRowIndex, newColIndex, _worksheet, _document);
        }

        /// <summary>
        /// If this Range is made up of columns/rows (as opposed to cells),
        /// returns the (index)th such column or row from the left/top. The index is
        /// 1-based so that 1 means the first column/row in the range.
        /// </summary>
        public Range Item(int index) {
            if (_type == RangeType.Columns) {
                return new Range(
                    _minRowIndex,
                    _minColIndex + index - 1,
                    _maxRowIndex,
                    _minColIndex + index - 1,
                    RangeType.Columns,
                    _worksheet,
                    _document
                );
            }
            if (_type == RangeType.Rows) {
                return new Range(
                    _minRowIndex + index - 1,
                    _minColIndex,
                    _minRowIndex + index - 1,
                    _maxColIndex,
                    RangeType.Rows,
                    _worksheet,
                    _document
                );
            }

            throw new ArgumentException("Can only call the Item method " +
                "with 1 argument when the Range type is rows or columns.");
        }
        
        /// <summary>
        /// Returns a Range object that represents the entire column (or columns) that
        /// contains the specified range.
        /// </summary>
        public Range EntireColumn() {
            return new Range(
                1,
                _minColIndex,
                ROW_INDEX_INFINITY,
                _maxColIndex,
                RangeType.Columns,
                _worksheet,
                _document
            );
        }

        /// <summary>
        /// Returns a Range object that represents the entire row (or rows) that
        /// contains the specified range.
        /// </summary>
        public Range EntireRow() {
            return new Range(
                _minRowIndex,
                1,
                _maxRowIndex,
                COL_INDEX_INFINITY,
                RangeType.Rows,
                _worksheet,
                _document
            );
        }

        /// <summary>
        /// Returns a Range object matching the specified range string, but relative to
        /// the top left corner of this Range. The Microsoft Excel equivalent of this
        /// method is Range.Range.
        /// </summary>
        /// <param name="rangeString">
        /// For example, "A2:IV65536" or "B3"
        /// </param>
        public Range RelativeRange(string rangeString) {
            (
                int relMinRowIndex, 
                int relMinColIndex, 
                int relMaxRowIndex, 
                int relMaxColIndex
            ) = CoordinateUtils.RangeStringToCoords(rangeString);
            return new Range(
                this._minRowIndex + relMinRowIndex - 1,
                this._minColIndex + relMinColIndex - 1,
                this._minRowIndex + relMaxRowIndex - 1,
                this._minColIndex + relMaxColIndex - 1,
                this._type,
                this._worksheet,
                this._document
            );
        }

        /// <summary>
        /// Returns a Range object that represents the cell at the end of the region that
        /// contains the source range. Equivalent to pressing END+UP ARROW, END+DOWN ARROW, 
        /// END+LEFT ARROW, or END+RIGHT ARROW.
        /// </summary>
        public Range End(XlDirection direction) {
            if (!this.IsOneByOne()) {
                throw new System.Exception("For now we only implement Range.End for " +
                    "1-by-1 ranges. Please expand the implementation if needed.");
            }

            Range firstCellOver;
            try {
                firstCellOver = this.AdjacentCell(direction);
            }
            catch(NoAdjacentCellException e) {
                // If this cell is an edge cell and we are moving towards the edge, we 
                // can't go anywhere, so we just return this cell.
                return this;
            }            

            // If this cell is nonempty *and* the first cell in the given direction is 
            // nonempty, then the End is the last nonempty cell in the given direction.
            if (!this.IsEmptyCell() && !firstCellOver.IsEmptyCell()) {
                Range curCell = this;
                Range nextCell = firstCellOver;
                while (!nextCell.IsEmptyCell()) {
                    
                    // Shift the cur cell over.
                    curCell = nextCell;

                    // If one of the cur cell indices is infinite, this is like hitting 
                    // the far edge of the spreadsheet, so we stop here.
                    if (curCell.MaxRowIndexIsInfinite() 
                            || curCell.MaxColIndexIsInfinite()) {
                        return curCell;
                    }                   

                    // Try to shift the next cell over, but we have to stop with the cur
                    // cell if we hit the edge.
                    try {
                        nextCell = nextCell.AdjacentCell(direction);
                    }
                    catch (NoAdjacentCellException) {
                        return curCell;
                    }                   
                }

                // When the next cell is empty, the current cell is the last nonempty
                // cell, which is what we are seeking.
                return curCell;
            }

            // Otherwise, the End is the first nonempty cell in the given direction
            // (not including this cell itself).
            else {
                Range curCell = firstCellOver;
                while (curCell.IsEmptyCell()) {
                    try {
                        curCell = curCell.AdjacentCell(direction);
                    }
                    catch (NoAdjacentCellException) {
                        // If we hit the edge, we stop here.
                        return curCell;
                    }

                    // If one of the cur cell indices is infinite, this is like hitting 
                    // the far edge of the spreadsheet, so we stop here.
                    if (curCell.MaxRowIndexIsInfinite()
                            || curCell.MaxColIndexIsInfinite()) {
                        return curCell;
                    }
                }

                return curCell;
            }
        }

        /// <summary>
        /// If this is a single cell, returns the address of that cell in R1C1 form.
        /// For example, cell B4 would give "R4C2". If this is not a single cell,
        /// throws an exception.
        /// </summary>
        public String Address() {
            if (!this.IsOneByOne()) {
                throw new Exception("Address of range is currently only implemented " +
                    "for 1-by-1 ranges.");
            }

            return $"R{_minRowIndex}C{_minColIndex}";
        }

        /// <summary>
        /// Sorts the values in the range.
        /// </summary>
        /// <param name="key">
        /// A 1-by-1 range. If we are sorting by rows, then the row of this range is the row
        /// that will end up sorted. If we are sorting by columns, then the column of this range
        /// is the column that will end up sorted.
        /// </param>
        /// <param name="order">
        /// Ascending or descending.
        /// </param>
        /// <param name="orientation">
        /// Whether we are sorting by rows or columns. Sorting by rows means that the columns
        /// remain internally the same but are rearranged relative to each other. Sorting by
        /// columns means that the rows remain internally the same but are rearranged relative
        /// to each other. 
        /// </param>
        public void Sort(
            Range key, 
            XlSortOrder order, 
            XlSortOrientation orientation = XlSortOrientation.xlSortRows
        ) {
            if (!key.IsOneByOne()) {
                throw new Exception("Sort is currently only implemented for 1-by-1 ranges.");
            }

            // Sorting activates the worksheet, if it wasn't already active.
            _worksheet.Activate();

            bool sortByColumn = orientation == XlSortOrientation.xlSortColumns;
            int keyIndex = sortByColumn ? key._minColIndex : key._minRowIndex;
            bool ascending = order == XlSortOrder.xlAscending;

            _document.Sort(
                _minRowIndex,
                _minColIndex,
                _maxRowIndex,
                _maxColIndex,
                sortByColumn,
                keyIndex,
                ascending
            );
        }

        /// <summary>
        /// Copies the contents of an ADO Recordset object onto a worksheet, beginning
        /// at the upper-left corner of the specified range.
        /// </summary>
        public void CopyFromRecordset(ADODB.Recordset recordset) {
            // Copying from recordset activates the worksheet, if it wasn't already active.
            _worksheet.Activate();

            // Loop through all the records in the set.
            for (int recordNum = 0; recordNum < recordset.RecordCount; recordNum++) {
                // Write all the fields in the record.
                for (int fieldNum = 0; fieldNum < recordset.Fields.Count; fieldNum++) {
                    writeRecordsetFieldToCell(
                        recordset.Fields[fieldNum],
                        _minRowIndex + recordNum,
                        _minColIndex + fieldNum
                    );
                }

                recordset.MoveNext();
            }
        }

        private void writeRecordsetFieldToCell(ADODB.Field field, int row, int col) {
            try {
                // Don't do anything if the field value is empty
                if (field.Value + "" != "") {
                    Range cellToWriteTo = new Range(row, col, _worksheet, _document);

                    switch (field.Type) {
                        case ADODB.DataTypeEnum.adInteger:
                            cellToWriteTo.Value = Int32.Parse(field.Value.ToString());
                            break;
                        case ADODB.DataTypeEnum.adSmallInt:
                            cellToWriteTo.Value = Int16.Parse(field.Value.ToString());
                            break;
                        case ADODB.DataTypeEnum.adVarChar:
                            cellToWriteTo.Value = field.Value.ToString();
                            break;
                        case ADODB.DataTypeEnum.adDBDate:
                            cellToWriteTo.Value = DateTime.Parse(field.Value.ToString());
                            break;
                        // The default case will probably throw an error, but info will be written 
                        // to the sheet so a new type can be added.
                        default:
                            _document.SetCellValue(row, col, (string) field.Value);
                            break;
                    }
                }
            }
            catch (Exception e) {
                _document.SetCellValue(row, col, "Error writing value (" + field.Value + ") of type : " 
                    + field.Type + ".  Exception: " + e.Message);
            }
        }

        /// <summary>
        /// Deletes the range, which must consist of entire rows and/or entire columns.
        /// </summary>
        public void Delete() {
            // Deleting a range from a worksheet activates that worksheet, if it 
            // wasn't already active.
            _worksheet.Activate();

            if (this.IsEntireColumns()) {
                int numColsToDelete = _maxColIndex - _minColIndex + 1;
                _document.DeleteColumn(_minColIndex, numColsToDelete);
            }
            else if (this.IsEntireRows()) {
                int numRowsToDelete = _maxRowIndex - _minRowIndex + 1;
                _document.DeleteRow(_minRowIndex, numRowsToDelete);
            }
            else {
                throw new Exception("Tried to delete Range that doesn't consist of " +
                "entire rows or entire columns.");
            }           
        }

        /// <summary>
        /// Activates a single cell.
        /// </summary>
        public void Activate() {
            if (!this.IsOneByOne()) {
                throw new Exception("Can only activate a 1-by-1 range.");
            }

            // Activating a cell activates the worksheet, if it wasn't already.
            _worksheet.Activate();

            _document.SetActiveCell(_minRowIndex, _minColIndex);
        }

        /// <summary>
        /// Changes the width of the columns in the range or the height of the rows 
        /// in the range to achieve the best fit.
        /// </summary>
        public void AutoFit() {
            // Autofitting columns activates this worksheet, if it wasn't already.
            _worksheet.Activate();

            if (_type == RangeType.Columns) {
                if (!this.IsEntireColumns()) {
                    throw new Exception("Cannot autofit partial columns.");
                }
                _document.AutoFitColumn(_minColIndex, _maxColIndex);
            }
            else if (_type == RangeType.Rows) {
                if (!this.IsEntireRows()) {
                    throw new Exception("Cannot autofit partial rows.");
                }
                _document.AutoFitRow(_minRowIndex, _maxRowIndex);
            }           
        }

        /// <summary>
        /// If this is a single cell, returns the cell adjacent to this one in
        /// the given direction. If this is not a single cell, throws an exception.
        /// If there is no adjacent cell in the given direction because we are at the
        /// edge of the sheet, throws an exception.
        /// </summary>
        private Range AdjacentCell(XlDirection direction) {
            if (!this.IsOneByOne()) {
                throw new System.Exception("Called AdjacentCell on Range containing " +
                    "multiple cells.");
            }            

            int newRowIndex = _minRowIndex;
            int newColIndex = _minColIndex;
            if (direction == XlDirection.xlToLeft) {
                newColIndex -= 1;
            }
            else if (direction == XlDirection.xlUp) {
                newRowIndex -= 1;
            }
            else if (direction == XlDirection.xlToRight) {
                newColIndex += 1;
            }
            else if (direction == XlDirection.xlDown) {
                newRowIndex += 1;
            }

            // If we have moved off the sheet, throw an exception.
            if (newRowIndex < 1 || newColIndex < 1) {
                throw new NoAdjacentCellException();
            }

            return new Range(
                newRowIndex,
                newColIndex,
                _worksheet,
                _document
            );
        }

        private bool IsEntireColumns() {
            return _minRowIndex == 1 && MaxRowIndexIsInfinite();
        }

        private bool IsEntireRows() {
            return _minColIndex == 1 && MaxColIndexIsInfinite();
        }

        private bool MaxColIndexIsInfinite() {
            return _maxColIndex >= COL_INDEX_INFINITY;
        }

        private bool MaxRowIndexIsInfinite() {
            return _maxRowIndex >= ROW_INDEX_INFINITY;
        }

        private bool IsOneByOne() {
            return _minRowIndex == _maxRowIndex && _minColIndex == _maxColIndex;
        }

        /// <summary>
        /// If this is a single cell, returns whether it is empty. If this is not a single 
        /// cell, throws an exception.
        /// </summary>
        private bool IsEmptyCell() {
            if (!this.IsOneByOne()) {
                throw new System.Exception("Called IsEmptyCell on Range containing " +
                    "multiple cells. That request is ambiguous.");
            }

            return this.Value == null;
        }

        /// <summary>
        /// A string describing the format of numbers in this range. For example,
        /// "0.00" for numeric with 2 decimal places or "@" for pure text.
        /// You can see all the options by opening Excel, selecting a cell, 
        /// and doing Format Cells > Number > Custom.
        /// </summary>
        public string NumberFormat {
            set {
                // Setting a number format on a worksheet activates that worksheet, 
                // if it wasn't already.
                _worksheet.Activate();

                // If this is entire columns, it would take forever to set all the cells
                // individually, so we set whole columns.
                if (this.IsEntireColumns()) {
                    for (int i = _minColIndex; i <= _maxColIndex; i++) {
                        SLStyle style = _document.GetColumnStyle(i);
                        style.FormatCode = value;
                        _document.SetColumnStyle(i, style);
                    }
                }
                // Otherwise, it's fine to set cells individually.
                else {
                    for (int row = _minRowIndex; row <= _maxRowIndex; row++) {
                        for (int col = _minColIndex; col <= _maxColIndex; col++) {
                            SLStyle style = _document.GetCellStyle(row, col);
                            style.FormatCode = value;
                            _document.SetCellStyle(row, col, style);
                        }
                    }
                }
            }            
        }

        /// <summary>
        /// The width of the columns in this range.
        /// </summary>
        public int ColumnWidth {
            set {
                // Setting a column width on a worksheet activates that worksheet, 
                // if it wasn't already.
                _worksheet.Activate();

                _document.SetColumnWidth(_minColIndex, _maxColIndex, value);
            }
        }

        /// <summary>
        /// The content of the range, which can be an int, double, date, or string. 
        /// Or if the cell is empty, getting the Value gives null. The value 
        /// can only be set and retrieved for ranges representing single cells.
        /// </summary>
        public object Value {
            set {
                // Setting a value on a worksheet activates that worksheet, if it
                // wasn't already.
                _worksheet.Activate();

                if (_type != RangeType.Cells) {
                    throw new Exception("Value of range can only be set " +
                        "for ranges representing cells.");
                }
                if (!IsOneByOne()) {
                    throw new Exception("Value of range can only be set " +
                        "for one-by-one ranges.");
                }

                if (value.GetType() == typeof(int)) {
                    _document.SetCellValue(_minRowIndex, _minColIndex, (int)value);
                }
                else if (value.GetType() == typeof(double)) {
                    _document.SetCellValue(_minRowIndex, _minColIndex, (double)value);
                }
                else if (value.GetType() == typeof(string)) {
                    _document.SetCellValue(_minRowIndex, _minColIndex, (string)value);
                }
                else if (value.GetType() == typeof(DateTime)) {
                    _document.SetCellValue(_minRowIndex, _minColIndex, (DateTime)value);
                    
                    SLStyle style = _document.GetCellStyle(_minRowIndex, _minColIndex);
                    style.FormatCode = DATE_FORMAT_CODE;
                    _document.SetCellStyle(_minRowIndex, _minColIndex, style);
                }
                else {
                    throw new Exception(
                        "Can only set cell value with int, double, string, or date.");
                }
            }
            get {
                if (_type != RangeType.Cells) {
                    throw new Exception("Value of range can only be retrieved " +
                        "for ranges representing cells.");
                }
                if (!IsOneByOne()) {
                    throw new Exception("Value of range can only be retrieved " +
                        "for one-by-one ranges.");
                }

                string prevActiveSheet = _document.GetCurrentWorksheetName();
                _worksheet.Activate();

                // Get everything we may need while we have the worksheet active.
                DateTime valueAsDateTime = _document.GetCellValueAsDateTime(_minRowIndex, _minColIndex);
                string valueAsString = _document.GetCellValueAsString(_minRowIndex, _minColIndex);
                string cellFormatCode = _document.GetCellStyle(_minRowIndex, _minColIndex).FormatCode;

                _document.SelectWorksheet(prevActiveSheet);

                if (valueAsString == "") {
                    // Return null to indicate that the cell is empty.
                    return null;
                }
                // We rely on the format code to check if the value is a date. (NOTE: If 
                // you someday find that dates are being read as numbers, you should add more 
                // date format codes here.)
                if (cellFormatCode == DATE_FORMAT_CODE) {
                    return valueAsDateTime;
                }
                if (Double.TryParse(valueAsString, out double doubleVal)) {
                    return doubleVal;
                }                
                return valueAsString;
            }            
        }

        /// <summary>
        /// If this range is a single cell, represents the inside of the cell,
        /// which can for example be given a fill color.
        /// </summary>
        public Interior Interior {
            get {
                if (_type != RangeType.Cells) {
                    throw new Exception("Interior of range can only be retrieved " +
                        "for ranges representing cells.");
                }
                if (!IsOneByOne()) {
                    throw new Exception("Interior of range can only be retrieved " +
                        "for one-by-one ranges.");
                }
                return new Interior(_minRowIndex, _minColIndex, _worksheet, _document);
            }            
        }

        /// <summary>
        /// Returns the number of the first row in the range.
        /// </summary>
        public int Row {
            get {
                return _minRowIndex;
            }
        }

        /// <summary>
        /// Returns the number of the first column in the range.
        /// </summary>
        public int Column {
            get {
                return _minColIndex;
            }
        }
    }
}
