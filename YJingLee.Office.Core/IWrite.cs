using System.Collections.Generic;

namespace YJingLee.Office.Core
{
    public interface IBasicWrite
    {
        void WriteValue(int sheetIndex, int rowIndex, int cellIndex, dynamic value, int styleIndex, string formula = null);
        void CreateSheet(string name);
        byte[] WriteStream();
        void WriteFile(string filePath);
    }

    public interface IWrite
    {
        void WriteTitle(int sheetIndex, int rowIndex, dynamic[] titles);
        void WriteProperty<T>(int sheetIndex, int rowIndex, T firstEntity, object secondEntity = null);
        void WriteEnumerable<T>(int sheetIndex, int rowIndex, IEnumerable<T> entities);
        void WriteObject<T>(int sheetIndex, int rowIndex, ICollection<T> entities);
    }
}