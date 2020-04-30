using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using SpreadsheetLight;

namespace ExcelWrapper {
    /// <summary>
    /// Provides utility functions to interpret strings like "A"
    /// (meaning column 1), "A1" (meaning cell (1, 1)) and "B2:LC123"
    /// (meaning a particular bounded region).
    /// </summary>
    public class CoordinateUtils {
        private static Regex RANGE_STRING_REGEX = new Regex(
            @"^[A-Z]+\d+(:[A-Z]+\d+)?$", RegexOptions.IgnoreCase);

        /// <summary>
        /// Converts a range string to an array of coordinates bounding 
        /// a region. Range strings have either the form "A1" or "A1:B2".
        /// For example, "A6:C10" => (6, 1, 10, 3).
        /// </summary>    
        public static (int minRowIndex, int minColIndex, int maxRowIndex, int maxColIndex) 
            RangeStringToCoords(string rangeString) 
        {
            if (!RANGE_STRING_REGEX.IsMatch(rangeString)) {
                throw new ArgumentException($"Invalid range reference: '{rangeString}'.");
            }

            string[] rangeCorners = rangeString.Contains(":") ?
                 rangeString.Split(new char[] { ':' }) :
                 new string[] { rangeString, rangeString };

            (int corner1Row, int corner1Col) = CellReferenceToCoords(rangeCorners[0]);
            (int corner2Row, int corner2Col) = CellReferenceToCoords(rangeCorners[1]);

            int minRowIndex = Math.Min(corner1Row, corner2Row);
            int maxRowIndex = Math.Max(corner1Row, corner2Row);
            int minColIndex = Math.Min(corner1Col, corner2Col);
            int maxColIndex = Math.Max(corner1Col, corner2Col);

            return (minRowIndex, minColIndex, maxRowIndex, maxColIndex);
        }

        /// <summary>
        /// Converts a cell reference such as "A1" or "KT14" to a row index and
        /// column index. For example, "B68" => (68, 2). 
        /// </summary>
        public static (int rowIndex, int colIndex) 
            CellReferenceToCoords(string cellString) 
        {
            int rowIndex, colIndex;
            if (SLDocument.WhatIsRowColumnIndex(cellString, out rowIndex, out colIndex)) {
                return (rowIndex, colIndex);
            }

            throw new ArgumentException($"Invalid cell reference: '{cellString}'.");
        }

        /// <summary>
        /// Converts either a column name or a cell reference to a column index.
        /// For example, "B" => 2; "AB" => 28; "C5" => 3.
        /// </summary>
        public static int ToColumnIndex(string colLetters) {
            return SLConvert.ToColumnIndex(colLetters);
        }
    }
}
