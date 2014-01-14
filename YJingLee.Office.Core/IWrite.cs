using System.Collections.Generic;

namespace YJingLee.Office.Core
{
    public interface IBasicWrite
    {
        void CreateRow(int sheetIndex, int rowIndex);
        void WriteValue(int sheetIndex, int rowIndex, int cellIndex, dynamic value, int styleIndex, string formula = null);
        void CreateSheet(string name);
        byte[] WriteStream();
        void WriteFile(string filePath);
    }

    public interface IWrite
    {
        void WriteTitle(string[] titles, int sheetIndex, int rowIndex, int cellIndex = 0, int styleIndex = 1);
        int WriteProperty<T>(T entity, int sheetIndex, int rowIndex, int cellIndex = 0, int styleIndex = 2);
        void WriteEnumerable<T>(IEnumerable<T> entities, int sheetIndex, int rowIndex);
        void WriteObject<T>(ICollection<T> entities, int sheetIndex, int rowIndex);
    }
}