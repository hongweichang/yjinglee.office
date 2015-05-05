using System;
using System.Collections.Generic;

namespace YJingLee.Office.Core
{
    public interface IRead
    {
        void Read(int sheetIndex, int rowIndex, Action<dynamic[]> action);
        T ReadProperty<T>(int sheetIndex, int rowIndex);
        IEnumerable<T> ReadEnumerable<T>(int sheetIndex, int rowIndex);
    }
}
