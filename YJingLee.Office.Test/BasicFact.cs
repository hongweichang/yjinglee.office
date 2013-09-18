using System;
using System.IO;
using Xunit;
using YJingLee.Office.Npoi;

namespace YJingLee.Office.Test
{
    public class BasicFact: AbstractFact
    {
        [Fact]
        public void BasicBlank()
        {
            var excel = new Excel();
            excel.CreateSheet("Test");
            excel.WriteObject(0, 0, Reports);
            excel.SetColumnWidth(0, 0, new[] { 5, 35 });
            excel.WriteFile(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "aa.xlsx"));
        }

        [Fact]
        public void BasicDefault()
        {
            var excel = new Excel(new DefaultStyle());
            excel.CreateSheet("Test");
            excel.WriteObject(0, 0, Reports);
            excel.SetColumnWidth(0, 0, new[] { 5, 35 });
            excel.WriteFile(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "aa.xlsx"));
        }

        [Fact]
        public void BasicBlue()
        {
            var excel = new Excel(new BlueStyle());
            excel.CreateSheet("Test");
            excel.WriteObject(0, 0, Reports);
            excel.SetColumnWidth(0, 0, new[] { 5, 35 });
            excel.WriteFile(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "aa.xlsx"));
        }
    }
}
