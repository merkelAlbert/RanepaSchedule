using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using RanepaSchedule.Extensions;
using RanepaSchedule.Models;
using Spire.Doc;
using Spire.Doc.Collections;


namespace RanepaSchedule.Services
{
    public class ScheduleService
    {
        private const string Url = "http://op.klg.ranepa.ru/sites/default/files/workers/Raspis/raspis4.doc";
        private const string Path = "schedule.doc";
        private const string Group = "ДГиМУ-15";
        private int _groupIndex;
        
        //получаем строки новых дней
        private List<int> GetNewDayIndexes(RowCollection rows)
        {
            var indexes = new List<int>();
            int dayCellIndex = 1;
            indexes.Add(0);
            for (int i = 2; i < rows.Count; i++)
            {
                var currentValue = rows[i].Cells[dayCellIndex].Paragraphs.GetText();

                if (currentValue.Contains("1"))
                    indexes.Add(i - 1);
            }

            indexes.Add(rows.Count - 1);
            return indexes;
        }

        public async Task<List<ScheduleModel>> GetSchedule()
        {
            using (var client = new HttpClient())
            {
                var response = await client.GetByteArrayAsync(Url);
                if (response.Length != 0)
                {
                    var schedule = new List<ScheduleModel>();

                    if (File.Exists(Path))
                        File.Delete(Path);

                    using (FileStream fs = File.Create(Path))
                    {
                        fs.Write(response, 0, response.Length);
                    }

                    var document = new Document(Path);
                    var section = document.Sections[0];
                    var table = section.Tables[0];
                    var rows = table.Rows;

                    //получаем нужный столбец
                    for (int i = 0; i < rows[0].Cells.Count; i++)
                    {
                        var cell = rows[0].Cells[i];
                        var group = cell.Paragraphs.GetText().Replace(" ", String.Empty).ToLower();
                        if (cell.Paragraphs.GetText().Replace(" ",String.Empty).ToLower() == Group.ToLower())
                        {
                            _groupIndex = i;
                            break;
                        }
                    }

                    var newDayIndexes = GetNewDayIndexes(rows);

                    //цикл по дням
                    for (int i = 0; i < newDayIndexes.Count - 1; i++)
                    {
                        var startIndex = newDayIndexes[i] + 1;
                        var endIndex = newDayIndexes[i + 1] + 1;

                        var scheduleOfDay = new ScheduleModel();
                        scheduleOfDay.DayOfWeek = rows[startIndex]
                            .Cells[0].Paragraphs.GetText();

                        //цикл по предметам в дне
                        for (int j = startIndex; j < endIndex; j++)
                        {
                            var subject = rows[j].Cells[_groupIndex].Paragraphs.GetText();
                            if (subject.IndexOf(')') != subject.Length - 1 && subject.IndexOf(')') != -1)
                            {
                                subject = subject.Insert(subject.IndexOf(')') + 1, "\n\n");
                            }

                            scheduleOfDay.Subjects.Add(subject);
                            scheduleOfDay.Times.Add(rows[j].Cells[2].Paragraphs.GetText());
                        }

                        schedule.Add(scheduleOfDay);
                    }

                    return schedule;
                }

                return null;
            }
        }
    }
}