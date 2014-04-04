using System.ComponentModel;

namespace YJingLee.Office.Test.Domain
{
    public class Report
    {
        [Description("标记")]
        public int Id { get; set; }
        [Description("名称")]
        public string Name { get; set; }
        [Description("数量")]
        public long Count { get; set; }
    }

    public class Data
    {
        public int Type { get; set; }
        public long Count1 { get; set; }
        public long Count2 { get; set; }
    }
}