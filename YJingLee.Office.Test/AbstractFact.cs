using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using YJingLee.Office.Test.Domain;

namespace YJingLee.Office.Test
{
    public abstract class AbstractFact
    {
        protected readonly ICollection<Report> Reports;
        protected readonly IDictionary<int, ICollection<Data>> Datas;

        protected AbstractFact()
        {
            Reports = new Collection<Report>();
            Datas = new Dictionary<int, ICollection<Data>>();
            for (var i = 0; i < 10; i++)
            {
                Reports.Add(new Report { Id = i, Name = Guid.NewGuid().ToString(), Count = i * 5 });
                var data = new List<Data>();
                for (int j = 0; j < 3; j++)
                {
                    data.Add(new Data { Type = j, Count1 = j * 2, Count2 = j * 3 });
                }
                Datas[i] = data;
            }
        }
    }
}