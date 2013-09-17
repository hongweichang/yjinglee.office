using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using YJingLee.Office.Npoi;

namespace YJingLee.Office.Test
{
//..\..\..\packages\ILRepack.1.22.2\tools\ILRepack.exe /out:YJingLee.Office.Temp.dll YJingLee.Office.Core.dll YJingLee.Office.Npoi.dll
//..\..\..\packages\ILRepack.1.22.2\tools\ILRepack.exe /internalize /out:YJingLee.Office.dll YJingLee.Office.Temp.dll NPOI.dll NPOI.OOXML.dll NPOI.OpenXml4Net.dll NPOI.OpenXmlFormats.dll ICSharpCode.SharpZipLib.dll
    public class Report
    {
        [Description("标记")]
        public int Id { get; set; }
        [Description("名称")]
        public string Name { get; set; }
    }
    class Program
    {
        private static readonly ICollection<Report> Reports = new Collection<Report>();

        private static void Main()
        {
            for (var i = 0; i < 10; i++)
            {
                Reports.Add(new Report {Id = i*100, Name = Guid.NewGuid().ToString()});
            }
            var excel = new Excel(new DefaultStyle());
            excel.CreateSheet("Test");
            excel.WriteObject(0, 0, Reports);
            excel.SetColumnWidth(0, 0, new [] {5, 35});
            excel.WriteFile(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "aa.xlsx"));
        }
    }
}
