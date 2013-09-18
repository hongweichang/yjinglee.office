using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using YJingLee.Office.Test.Domain;

namespace YJingLee.Office.Test
{
    public abstract class AbstractFact
    {
        protected readonly ICollection<Report> Reports;

        protected AbstractFact()
        {
            Reports = new Collection<Report>();
            for (var i = 0; i < 10; i++)
            {
                Reports.Add(new Report { Id = i * 100, Name = Guid.NewGuid().ToString() });
            }
        }
    }
}
