namespace ExcelWrapper {
    /// <summary>
    /// A direction to move within a worksheet.
    /// </summary>
    public enum XlDirection {
        xlDown = -4121,
        xlToLeft = -4159,
        xlToRight = -4161,
        xlUp = -4162
    }

    public static class XlDirectionExtensions {
        /// <summary>
        /// Returns the direction as a tuple representing the row and column
        /// index shift. For example, xlUp.toRowColShift() => (-1, 0).
        /// </summary>
        public static (int, int) toRowColShift(this XlDirection direction) {
            switch(direction) {
                case XlDirection.xlDown: return (1, 0);
                case XlDirection.xlToLeft: return (0, -1);
                case XlDirection.xlToRight: return (0, 1);
                case XlDirection.xlUp: return (-1, 0);
                default:
                    throw new System.Exception("XlDirection not recognized.");
            }
        }
    }
}