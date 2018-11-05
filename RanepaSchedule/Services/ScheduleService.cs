using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using RanepaSchedule.Models;

namespace RanepaSchedule.Services
{
    public class ScheduleService
    {
        private const string Url = "http://op.klg.ranepa.ru/sites/default/files/workers/Raspis/raspis4.doc";
        private const string Path = "schedule.doc";
        private const string Group = "ДГиМУ-15";
        private int _groupIndex;

        //получаем строки новых дней
        private List<int> GetNewDayIndexes(List<TableRow> rows)
        {
            var indexes = new List<int>();
            int dayCellIndex = 1;
            indexes.Add(0);

            var prevValue = rows[1].Elements<TableCell>().ElementAt(dayCellIndex).InnerText;
            for (int i = 2; i < rows.Count; i++)
            {
                var currentValue = rows[i].Elements<TableCell>().ElementAt(dayCellIndex).InnerText;
                if (int.Parse(currentValue) < int.Parse(prevValue))
                    indexes.Add(i - 1);
                prevValue = currentValue;
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

                    using (var document = WordprocessingDocument.Open(Path, false))
                    {
                        var body = document.MainDocumentPart.Document.Body;
                        var table = body.Elements<Table>().First();
                        var rows = table.Elements<TableRow>().ToList();

                        //получаем нужный столбец
                        for (int i = 0; i < rows[0].Elements<TableCell>().ToList().Count; i++)
                        {
                            var cell = rows[0].ElementAt(i);
                            if (cell.InnerText.ToLower() == Group.ToLower())
                            {
                                _groupIndex = i - 1;
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
                                .Elements<TableCell>()
                                .ElementAt(0).InnerText;

                            //цикл по предметам в дне
                            for (int j = startIndex; j < endIndex; j++)
                            {
                                var subject = rows[j].Elements<TableCell>().ElementAt(_groupIndex)
                                    .InnerText;
                                if (subject.IndexOf(')') != subject.Length - 1 && subject.IndexOf(')') != -1)
                                {
                                    subject = subject.Insert(subject.IndexOf(')') + 1, "\n\n");
                                }

                                scheduleOfDay.Subjects.Add(subject);
                                scheduleOfDay.Times.Add(rows[j].Elements<TableCell>().ElementAt(2).InnerText);
                            }

                            schedule.Add(scheduleOfDay);
                        }
                    }

                    return schedule;
                }

                return null;
            }
        }
    }
}