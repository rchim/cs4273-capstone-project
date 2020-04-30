using System;
using System.Runtime.Serialization;

namespace ExcelWrapper {
    [Serializable]
    internal class NoAdjacentCellException : Exception {
        public NoAdjacentCellException() {
        }
    }
}