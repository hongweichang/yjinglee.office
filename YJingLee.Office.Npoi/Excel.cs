using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.Util;
using YJingLee.Office.Core;

namespace YJingLee.Office.Npoi
{
    public class Excel : IRead, IBasicWrite, IWrite, IStyle
    {
        private readonly InternalExcel _internalExcel;

        public Excel(IExcelStyle excelStyle = null, string filePath = null)
        {
            _internalExcel = new InternalExcel(excelStyle, filePath);
        }

        public void Read(int sheetIndex, int rowIndex, Action<dynamic[]> action)
        {
            var currentSheet = _internalExcel.GetWorkbook().GetSheetAt(sheetIndex);

            for (var i = rowIndex; i < currentSheet.LastRowNum; i++)
            {
                var currentRow = currentSheet.GetRow(i);
                var count = currentRow.LastCellNum;
                var value = new dynamic[count];
                for (var j = 0; j < count; j++)
                {
                    value[j] = currentRow.GetRowData(j);
                }
                action(value);
            }
        }


        public T ReadProperty<T>(int sheetIndex, int rowIndex)
        {
            var cellIndex = 0;
            var currentSheet = _internalExcel.GetWorkbook().GetSheetAt(sheetIndex);
            var currentRow = currentSheet.GetRow(rowIndex);

            var entity = Activator.CreateInstance(typeof(T));
            var properties = typeof(T).GetProperties();
            foreach (var propertyInfo in properties)
            {
                var value = currentRow.GetCell(cellIndex).GetCellData();
                propertyInfo.SetValue(entity, value, null);
                cellIndex++;
            }
            return (T)entity;
        }

        public IEnumerable<T> ReadEnumerable<T>(int sheetIndex, int rowIndex)
        {
            var currentSheet = _internalExcel.GetWorkbook().GetSheetAt(sheetIndex);
            ICollection<T> results = new List<T>(currentSheet.LastRowNum - rowIndex);

            for (var i = rowIndex; i < currentSheet.LastRowNum; i++)
            {
                results.Add(ReadProperty<T>(sheetIndex, i));
            }
            return results;
        }

        public void CreateRow(int sheetIndex, int rowIndex)
        {
            _internalExcel.CreateRow(sheetIndex, rowIndex);
        }

        public void WriteValue(int sheetIndex, int rowIndex, int cellIndex, dynamic value, int styleIndex, string formula = null)
        {
            _internalExcel.WriteValue(sheetIndex, rowIndex, cellIndex, value, styleIndex, formula);
        }

        public void CreateSheet(string name)
        {
            _internalExcel.CreateSheet(name);
        }

        public byte[] WriteStream()
        {
            return _internalExcel.WriteStream();
        }

        public void WriteFile(string filePath)
        {
            _internalExcel.WriteFile(filePath);
        }

        public void WriteTitle(string[] titles, int sheetIndex, int rowIndex, int cellIndex = 0, int styleIndex = 1)
        {
            CreateRow(sheetIndex, rowIndex);
            for (var i = 0; i < titles.Length; i++)
            {
                WriteValue(sheetIndex, rowIndex, cellIndex + i, titles[i], styleIndex);
            }
        }

        public int WriteProperty<T>(T entity, int sheetIndex, int rowIndex, int cellIndex = 0, int styleIndex = 2)
        {
            var thisIndex = cellIndex;
            var firstProperties = entity.GetProperties();
            foreach (var property in firstProperties)
            {
                var value = entity.GetValue(property);
                WriteValue(sheetIndex, rowIndex, thisIndex, value, styleIndex);
                thisIndex++;
            }
            return thisIndex;
        }

        public void WriteEnumerable<T>(IEnumerable<T> entities, int sheetIndex, int rowIndex)
        {
            foreach (var entity in entities)
            {
                CreateRow(sheetIndex, rowIndex);
                WriteProperty(entity, sheetIndex, rowIndex);
                rowIndex++;
            }
        }

        public void WriteObject<T>(ICollection<T> entities, int sheetIndex, int rowIndex)
        {
            if (!entities.Any())
                return;
            var titles = typeof(T).GetProperties().Select(o => o.GetDescription()).ToArray();

            WriteTitle(titles, sheetIndex, rowIndex);
            rowIndex++;
            WriteEnumerable(entities, sheetIndex, rowIndex);
        }

        public void SetColumnWidth(int sheetIndex, int firstColumn, int[] widths)
        {
            var currentSheet = _internalExcel.GetWorkbook().GetSheetAt(sheetIndex);
            for (var i = 0; i < widths.Length; i++)
            {
                currentSheet.SetColumnWidth(firstColumn + i, (widths[i] + 2)*256);
            }
        }

        public void SetStyle(int sheetIndex, int firstRow, int lastRow, int firstColumn, int lastColumn, int styleIndex)
        {
            var currentSheet = _internalExcel.GetWorkbook().GetSheetAt(sheetIndex);

            if (styleIndex == 0)
            {
                currentSheet.AddMergedRegion(new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn));
            }
            else
            {
                var cellStyle = _internalExcel.GetWorkbook().GetCellStyleAt((short) styleIndex);
                for (var i = firstRow; i <= lastRow; i++)
                {
                    for (var j = firstColumn; j <= lastColumn; j++)
                    {
                        currentSheet.GetRow(i).GetCell(j).CellStyle = cellStyle;
                    }
                }
            }
        }
    }
}