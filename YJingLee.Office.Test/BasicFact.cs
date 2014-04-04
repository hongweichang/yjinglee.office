using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using YJingLee.Office.Npoi;
using YJingLee.Office.Test.Domain;

namespace YJingLee.Office.Test
{
    public class BasicFact : AbstractFact
    {
        [Fact]
        public void BasicBlank()
        {
            var excel = new Excel();
            excel.CreateSheet("Test");
            excel.WriteObject(Reports, 0, 0);
            excel.SetColumnWidth(0, 0, new[] { 5, 35 });
            excel.WriteFile(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "aa.xlsx"));
        }

        [Fact]
        public void BasicDefault()
        {
            var excel = new Excel(new DefaultStyle());
            excel.CreateSheet("Test");
            excel.WriteObject(Reports, 0, 0);
            excel.SetColumnWidth(0, 0, new[] { 5, 35 });
            excel.WriteFile(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "aa.xlsx"));
        }

        [Fact]
        public void BasicBlue()
        {
            var excel = new Excel(new BlueStyle());
            excel.CreateSheet("Test");
            excel.WriteObject(Reports, 0, 0);
            excel.SetColumnWidth(0, 0, new[] { 5, 35 });
            excel.WriteFile(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "aa.xlsx"));
        }

        [Fact]
        public void Custom()
        {
            var excel = new Excel(new FormatStyle());//使用特殊样式类
            excel.CreateSheet("Sheet名称");//注意Sheet名称不能重复
            excel.WriteTitle(new[] { "Id", "名称", "个数", "比例" }, 0, 0);//填充标题栏目
            var rowIndex = 1;
            foreach (var entity in Reports)
            {
                excel.CreateRow(0, rowIndex);
                excel.WriteValue(0, rowIndex, 0, entity.Id, 2);
                excel.WriteValue(0, rowIndex, 1, entity.Name, 2);
                excel.WriteValue(0, rowIndex, 2, entity.Count, 2);
                excel.WriteValue(0, rowIndex, 3, null, 3, string.Format("$C{0}/$C{1}", rowIndex + 1, Reports.Count + 2));//填充公式并设置百分号样式
                rowIndex++;
            }
            excel.CreateRow(0, rowIndex);
            excel.WriteValue(0, rowIndex, 0, "总计", 2);
            excel.WriteValue(0, rowIndex, 1, "", 2);
            excel.WriteValue(0, rowIndex, 2, null, 2, string.Format("SUM(C{0}:C{1})", 2, rowIndex));
            excel.WriteValue(0, rowIndex, 3, null, 3, string.Format("SUM(D{0}:D{1})", 2, rowIndex));//这里注意styleIndex指定为3
            excel.SetColumnWidth(0, 0, new[] { 10, 35 });//设置Id和Name列宽

            excel.CreateSheet("详情");
            excel.WriteTitle(new[] { "Id", "名称", "个数", "类型", "统计1", "统计2", "公式1" }, 1, 0);
            rowIndex = 1;
            foreach (var entity in Reports)
            {
                ICollection<Data> entities;
                if (Datas.TryGetValue(entity.Id, out entities))
                {
                    foreach (var data in entities)
                    {
                        excel.CreateRow(1, rowIndex);
                        var cellIndex = excel.WriteProperty(entity, 1, rowIndex);//自动填空entity对象属性
                        cellIndex = excel.WriteProperty(data, 1, rowIndex, cellIndex);//在cellIndex后自动填空data对象属性
                        excel.WriteValue(1, rowIndex, cellIndex, null, 3, string.Format("IF($E{0}=0,0,$F{0}/$E{0})", rowIndex + 1));//填充公式并设置百分号样式
                        rowIndex++;
                    }
                    //设置样式,这里合并单元格
                    excel.SetStyle(1, rowIndex - entities.Count, rowIndex - 1, 0, 0, 0);
                    excel.SetStyle(1, rowIndex - entities.Count, rowIndex - 1, 1, 1, 0);
                    excel.SetStyle(1, rowIndex - entities.Count, rowIndex - 1, 2, 2, 0);
                }
            }
            excel.SetColumnWidth(1, 1, new[] { 35 });//设置Name列宽
            excel.WriteFile(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "aa.xlsx"));
        }
    }
}