using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace ExcelWrapper {
    class ColorUtils {
        /// <summary>
        /// Converts an Excel ColorIndex (an integer from 1 to 56) to a 
        /// System.Drawing.Color object representing the same color.
        /// </summary>
        public static Color excelColorIndexToColor(int index) {
            switch (index) {
                case 15:
                    return Color.FromArgb(192, 192, 192);
                case 36:
                    return Color.FromArgb(255, 255, 153);
                default:
                    throw new System.Exception($"Excel ColorIndex {index} not recognized.");
            }
        }
    }
}
