using System.Collections.Generic;

namespace YJingLee.Office.Core
{
    public interface IRead
    {
        T ReadProperty<T>(int sheetIndex, int rowIndex);
        IEnumerable<T> ReadEnumerable<T>(int sheetIndex, int rowIndex);
    }
}
