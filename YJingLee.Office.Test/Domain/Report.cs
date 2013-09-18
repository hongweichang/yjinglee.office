using System.ComponentModel;

namespace YJingLee.Office.Test.Domain
{
    public class Report
    {
        [Description("标记")]
        public int Id { get; set; }
        [Description("名称")]
        public string Name { get; set; }
    }
}