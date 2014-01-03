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

        public void WriteTitle(int sheetIndex, int rowIndex, dynamic[] titles)
        {
            CreateRow(sheetIndex, rowIndex);
            for (var i = 0; i < titles.Length; i++)
            {
                WriteValue(sheetIndex, rowIndex, i, titles[i], 1);
            }
        }

        public void WriteProperty<T>(int sheetIndex, int rowIndex, T firstEntity, object secondEntity = null)
        {
            CreateRow(sheetIndex, rowIndex);
            var cellIndex = 0;

            var firstProperties = firstEntity.GetProperties();
            foreach (var property in firstProperties)
            {
                var value = firstEntity.GetValue(property);
                WriteValue(sheetIndex, rowIndex, cellIndex, value, 2);
                cellIndex++;
            }
            var secondProperties = secondEntity.GetProperties();
            foreach (var property in secondProperties)
            {
                var value = secondEntity.GetValue(property);
                WriteValue(sheetIndex, rowIndex, cellIndex, value, 2);
                cellIndex++;
            }
        }

        public void WriteEnumerable<T>(int sheetIndex, int rowIndex, IEnumerable<T> entities)
        {
            foreach (var entity in entities)
            {
                WriteProperty(sheetIndex, rowIndex, entity);
                rowIndex++;
            }
        }

        public void WriteObject<T>(int sheetIndex, int rowIndex, ICollection<T> entities)
        {
            if (!entities.Any())
                return;
            dynamic[] titles = typeof(T).GetProperties().Select(o => o.GetDescription()).ToArray();

            WriteTitle(sheetIndex, rowIndex, titles);
            rowIndex++;
            WriteEnumerable(sheetIndex, rowIndex, entities);
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

            if (sheetIndex == 0)
            {
                currentSheet.AddMergedRegion(new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn));
            }
            else
            {
                for (var i = firstRow; i <= lastRow; i++)
                {
                    for (var j = firstColumn; j <= lastColumn; j++)
                    {
                        currentSheet.GetRow(i).GetCell(j).CellStyle = _internalExcel.GetWorkbook().GetCellStyleAt((short)sheetIndex);
                    }
                }
            }
        }
    }
}