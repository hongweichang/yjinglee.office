using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace YJingLee.Office.Npoi
{
    public class DefaultStyle : IExcelStyle
    {
        public void InitCellStyle(IWorkbook workbook)
        {
            var font = workbook.CreateFont();
            var title = workbook.CreateCellStyle();
            var content = workbook.CreateCellStyle();

            RegisterFont(font);
            RegisterTitleStyle(workbook, title);
            RegisterContentStyle(workbook, content);
        }

        public virtual void RegisterFont(IFont font)
        {
            font.SetFontHeight();
        }

        public virtual void RegisterTitleStyle(IWorkbook workbook, ICellStyle cellStyle)
        {
            cellStyle.SetBorder();
            cellStyle.SetBackgroundColor(HSSFColor.Red.Index);
            cellStyle.SetFont(workbook.GetFontAt(1));
        }

        public virtual void RegisterContentStyle(IWorkbook workbook, ICellStyle cellStyle)
        {
            cellStyle.SetBorder();
        }

        public virtual void RegisterCustomStyle(IWorkbook workbook)
        {

        }
    }

    public class BlueStyle : DefaultStyle
    {
        public override void RegisterTitleStyle(IWorkbook workbook, ICellStyle cellStyle)
        {
            base.RegisterTitleStyle(workbook, cellStyle);
            cellStyle.SetBackgroundColor(HSSFColor.Blue.Index);
        }
    }
}