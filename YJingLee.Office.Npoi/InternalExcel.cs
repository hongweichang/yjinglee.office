using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using YJingLee.Office.Core;

namespace YJingLee.Office.Npoi
{
    public interface IInternalExcel : IBasicWrite
    {
        IWorkbook GetWorkbook();
    }

    public class InternalExcel : IInternalExcel
    {
        private IWorkbook _workbook;

        public InternalExcel(IExcelStyle excelStyle = null, string filePath = null)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                _workbook = new XSSFWorkbook();
            }
            else
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    _workbook = WorkbookFactory.Create(fs);
                }
            }
            if (excelStyle != null)
                ExcelStyleUtil.RegisterStyle(excelStyle, _workbook);
        }

        public void CreateRow(int sheetIndex, int rowIndex)
        {
            _workbook.GetSheetAt(sheetIndex).CreateRow(rowIndex);
        }

        public void WriteValue(int sheetIndex, int rowIndex, int cellIndex, dynamic value, int styleIndex, string formula =null)
        {
            var currentCell = _workbook.GetSheetAt(sheetIndex).GetRow(rowIndex).CreateCell(cellIndex);
            if (value != null)
            {
                if (value is decimal || value is long || value is ulong)
                    currentCell.SetCellValue((double) value);
                else
                    currentCell.SetCellValue(value);
            }
            if (styleIndex != 0 && styleIndex < _workbook.NumCellStyles)
                currentCell.CellStyle = _workbook.GetCellStyleAt((short) styleIndex);
            if (!string.IsNullOrEmpty(formula))
                currentCell.CellFormula = formula;
        }

        public void CreateSheet(string name)
        {
            _workbook.CreateSheet(name);
        }

        public byte[] WriteStream()
        {
            var ms = new MemoryStream();
            _workbook.Write(ms);
            _workbook = null;
            return ms.ToArray();
        }

        public void WriteFile(string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                _workbook.Write(fs);
            }
            _workbook = null;
        }

        public IWorkbook GetWorkbook()
        {
            return _workbook;
        }
    }
}
