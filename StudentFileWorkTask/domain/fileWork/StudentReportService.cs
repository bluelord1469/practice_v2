using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
using StudentFileWorkTask.data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace TestReporter.domain.fileWork
{
    /// <summary>
    /// Формирование отчётов по студенту и по группе в PDF и Excel.
    /// </summary>
    internal class StudentReportService
    {
        private readonly IEnumerable<StudentResult> _allResults;

        public StudentReportService(IEnumerable<StudentResult> allResults)
        {
            _allResults = allResults ?? Enumerable.Empty<StudentResult>();
        }

        public List<string> GetStudentSuggestions(string query)
        {
            var all = _allResults
                .Select(r => r.Student)
                .Where(s => s != null)
                .Select(s => $"{s.Surname} {s.Name} {s.Patronymic}")
                .Distinct()
                .OrderBy(fio => fio)
                .ToList();

            if (string.IsNullOrWhiteSpace(query))
                return all;

            query = query.ToLower();
            return all.Where(fio => fio.ToLower().Contains(query)).ToList();
        }

        public List<string> GetGroupSuggestions(string query)
        {
            var all = _allResults
                .Select(r => r.Student?.Group?.GroupName)
                .Where(g => !string.IsNullOrWhiteSpace(g))
                .Distinct()
                .OrderBy(g => g)
                .ToList();

            if (string.IsNullOrWhiteSpace(query))
                return all;

            query = query.Trim().ToLower();
            return all.Where(g => g.ToLower().Contains(query)).ToList();
        }

        public void ExportStudentPdf(string fio)
        {
            var results = FilterByStudent(fio);
            if (!ValidateResults(results, fio, "студенту")) return;

            var dialog = new SaveFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                FileName = $"Отчет_{SanitizeFileName(fio)}_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (dialog.ShowDialog() != true) return;

            try
            {
                BuildStudentPdf(dialog.FileName, fio, results);
                MessageBox.Show($"PDF отчёт сохранён:\n{dialog.FileName}", "Готово",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка формирования PDF: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void ExportStudentExcel(string fio)
        {
            var results = FilterByStudent(fio);
            if (!ValidateResults(results, fio, "студенту")) return;

            var dialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = $"Отчет_{SanitizeFileName(fio)}_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (dialog.ShowDialog() != true) return;

            try
            {
                BuildStudentExcel(dialog.FileName, fio, results);
                MessageBox.Show($"Отчёт успешно создан!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка формирования Excel: {ex.Message}");
            }
        }

        public void ExportGroupPdf(string groupName)
        {
            var results = FilterByGroup(groupName);
            if (!ValidateResults(results, groupName, "группе")) return;

            var dialog = new SaveFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                FileName = $"Отчет_группа_{SanitizeFileName(groupName)}_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (dialog.ShowDialog() != true) return;

            try
            {
                BuildGroupPdf(dialog.FileName, groupName, results);
                MessageBox.Show($"PDF успешно создан!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка формирования PDF: {ex.Message}");
            }
        }

        public void ExportGroupExcel(string groupName)
        {
            var results = FilterByGroup(groupName);
            if (!ValidateResults(results, groupName, "группе")) return;

            var dialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = $"Отчет_группа_{SanitizeFileName(groupName)}_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (dialog.ShowDialog() != true) return;

            try
            {
                BuildGroupExcel(dialog.FileName, groupName, results);
                MessageBox.Show($"Отчёт успешно создан!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка формирования Excel: {ex.Message}");
            }
        }

        private bool ValidateResults(List<StudentResult> results, string key, string target)
        {
            if (results.Count == 0)
            {
                MessageBox.Show($"Не найдено результатов для {target} «{key}».");
                return false;
            }
            return true;
        }

        private List<StudentResult> FilterByStudent(string fio)
        {
            var normalized = NormalizeFio(fio);
            return _allResults
                .Where(r =>
                {
                    var s = r.Student;
                    if (s == null) return false;
                    var candidate = NormalizeFio($"{s.Surname} {s.Name} {s.Patronymic}");
                    return candidate == normalized || candidate.Contains(normalized);
                }).ToList();
        }

        private List<StudentResult> FilterByGroup(string groupName)
        {
            return _allResults
                .Where(r => string.Equals(r.Student?.Group?.GroupName?.Trim(),
                                          groupName?.Trim(),
                                          StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private static string NormalizeFio(string s) =>
            string.Join(" ", (s ?? "").Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                  .ToLower();

        private static string SanitizeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            return new string(name.Where(c => !invalid.Contains(c)).ToArray()).Replace(' ', '_');
        }

        private class TestAttemptRow
        {
            public int TestNumber { get; set; }
            public DateOnly? Date { get; set; }
            public double Score { get; set; }
            public double MaxScore { get; set; }
            public double Percent => MaxScore > 0 ? Score / MaxScore : 0;
        }

        private class GroupTestRow
        {
            public int TestNumber { get; set; }
            public string Theme { get; set; }
            public int StudentsCount { get; set; }
            public int AttemptsCount { get; set; }
            public double AvgScore { get; set; }
        }

        private List<TestAttemptRow> BuildStudentRows(List<StudentResult> studentResults)
        {
            var maxByTheme = ComputeMaxScoresByTheme();

            var attempts = studentResults
                .GroupBy(r => new
                {
                    Theme = r.Question?.Theme?.ThemeName ?? "Без темы",
                    r.Date
                })
                .Select(g => new
                {
                    g.Key.Theme,
                    g.Key.Date,
                    Score = g.Sum(x => x.Score)
                })
                .OrderBy(a => a.Theme)
                .ThenBy(a => a.Date)
                .ToList();

            var themeNumbers = attempts
                .Select(a => a.Theme)
                .Distinct()
                .OrderBy(t => t)
                .Select((t, i) => new { t, num = i + 1 })
                .ToDictionary(x => x.t, x => x.num);

            return attempts.Select(a => new TestAttemptRow
            {
                TestNumber = themeNumbers[a.Theme],
                Date = a.Date,
                Score = a.Score,
                MaxScore = maxByTheme.TryGetValue(a.Theme, out var mx) ? mx : 0
            }).ToList();
        }

        private Dictionary<string, double> ComputeMaxScoresByTheme()
        {
            var attempts = _allResults
                .GroupBy(r => new
                {
                    Theme = r.Question?.Theme?.ThemeName ?? "Без темы",
                    Student = r.Student,
                    r.Date
                })
                .Select(g => new
                {
                    g.Key.Theme,
                    Sum = g.Sum(x => x.Score)
                });

            return attempts
                .GroupBy(x => x.Theme)
                .ToDictionary(g => g.Key, g => g.Max(x => x.Sum));
        }

        private List<GroupTestRow> BuildGroupRows(List<StudentResult> groupResults)
        {
            var attempts = groupResults
                .GroupBy(r => new
                {
                    Theme = r.Question?.Theme?.ThemeName ?? "Без темы",
                    Student = r.Student,
                    r.Date
                })
                .Select(g => new
                {
                    g.Key.Theme,
                    Student = g.Key.Student,
                    Sum = g.Sum(x => x.Score)
                })
                .ToList();

            var themeNumbers = attempts
                .Select(a => a.Theme)
                .Distinct()
                .OrderBy(t => t)
                .Select((t, i) => new { t, num = i + 1 })
                .ToDictionary(x => x.t, x => x.num);

            return attempts
                .GroupBy(a => a.Theme)
                .Select(g => new GroupTestRow
                {
                    TestNumber = themeNumbers[g.Key],
                    Theme = g.Key,
                    StudentsCount = g.Select(x => x.Student).Distinct().Count(),
                    AttemptsCount = g.Count(),
                    AvgScore = g.Average(x => x.Sum)
                })
                .OrderBy(r => r.TestNumber)
                .ToList();
        }

        private void BuildStudentPdf(string path, string fio, List<StudentResult> studentResults)
        {
            var rows = BuildStudentRows(studentResults);
            var student = studentResults.First().Student;
            string groupName = student?.Group?.GroupName ?? "—";

            using var fs = new FileStream(path, FileMode.Create);
            var doc = new Document(PageSize.A4, 40, 40, 40, 40);
            PdfWriter.GetInstance(doc, fs);
            doc.Open();

            var baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf",
                                                BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            var titleFont = new Font(baseFont, 16, Font.BOLD);
            var labelFont = new Font(baseFont, 12, Font.BOLD);
            var valueFont = new Font(baseFont, 12, Font.NORMAL);
            var headerFont = new Font(baseFont, 11, Font.BOLD);
            var cellFont = new Font(baseFont, 10, Font.NORMAL);

            var title = new Paragraph("Отчёт по результатам тестирования", titleFont)
            { Alignment = Element.ALIGN_CENTER };
            doc.Add(title);
            doc.Add(new Paragraph("\n"));

            doc.Add(new Paragraph($"ФИО: {fio}", labelFont));
            doc.Add(new Paragraph($"Номер группы: {groupName}", valueFont));
            doc.Add(new Paragraph($"Количество пройденных тестов: {rows.Count}", valueFont));
            doc.Add(new Paragraph("\n"));

            var table = new PdfPTable(5) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 10f, 20f, 15f, 25f, 20f });

            string[] headers =
            {
                "№ теста", "Дата сдачи", "Результат теста",
                "Максимальное количество баллов", "Процент верных ответов"
            };

            foreach (var h in headers)
            {
                var phrase = new Phrase(h, headerFont) { Font = { Color = BaseColor.WHITE } };
                table.AddCell(new PdfPCell(phrase)
                {
                    BackgroundColor = new BaseColor(129, 166, 198),
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Padding = 5
                });
            }

            foreach (var r in rows)
            {
                table.AddCell(new PdfPCell(new Phrase(r.TestNumber.ToString(), cellFont))
                { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 4 });
                table.AddCell(new PdfPCell(new Phrase(r.Date?.ToString("dd.MM.yyyy") ?? "—", cellFont))
                { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 4 });
                table.AddCell(new PdfPCell(new Phrase(r.Score.ToString("0.##"), cellFont))
                { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 4 });
                table.AddCell(new PdfPCell(new Phrase(r.MaxScore.ToString("0.##"), cellFont))
                { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 4 });
                table.AddCell(new PdfPCell(new Phrase(r.Percent.ToString("0.00"), cellFont))
                { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 4 });
            }

            doc.Add(table);
            doc.Close();
        }

        private void BuildGroupPdf(string path, string groupName, List<StudentResult> groupResults)
        {
            var rows = BuildGroupRows(groupResults);

            using var fs = new FileStream(path, FileMode.Create);
            var doc = new Document(PageSize.A4, 40, 40, 40, 40);
            PdfWriter.GetInstance(doc, fs);
            doc.Open();

            var baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf",
                                                BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            var titleFont = new Font(baseFont, 16, Font.BOLD);
            var headerFont = new Font(baseFont, 11, Font.BOLD);
            var cellFont = new Font(baseFont, 10, Font.NORMAL);

            var title = new Paragraph($"Отчёт по группе {groupName}", titleFont)
            { Alignment = Element.ALIGN_CENTER };
            doc.Add(title);
            doc.Add(new Paragraph("\n"));

            var table = new PdfPTable(5) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 12f, 18f, 22f, 22f, 26f });

            string[] headers =
            {
                "№ теста", "Тема",
                "Сколько людей сдавали тест",
                "Сколько было попыток сдать тест",
                "Среднее количество баллов по всем попыткам"
            };

            foreach (var h in headers)
            {
                var phrase = new Phrase(h, headerFont) { Font = { Color = BaseColor.WHITE } };
                table.AddCell(new PdfPCell(phrase)
                {
                    BackgroundColor = new BaseColor(129, 166, 198),
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Padding = 5
                });
            }

            foreach (var r in rows)
            {
                table.AddCell(new PdfPCell(new Phrase(r.TestNumber.ToString(), cellFont))
                { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 4 });
                table.AddCell(new PdfPCell(new Phrase(r.Theme, cellFont)) { Padding = 4 });
                table.AddCell(new PdfPCell(new Phrase(r.StudentsCount.ToString(), cellFont))
                { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 4 });
                table.AddCell(new PdfPCell(new Phrase(r.AttemptsCount.ToString(), cellFont))
                { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 4 });
                table.AddCell(new PdfPCell(new Phrase(r.AvgScore.ToString("0.00"), cellFont))
                { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 4 });
            }

            doc.Add(table);
            doc.Close();
        }

        private void BuildStudentExcel(string path, string fio, List<StudentResult> studentResults)
        {
            var rows = BuildStudentRows(studentResults);
            var student = studentResults.First().Student;
            string groupName = student?.Group?.GroupName ?? "—";

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Отчёт по студенту");

            ws.Cell(1, 1).Value = "Отчёт по результатам тестирования";
            ws.Range(1, 1, 1, 5).Merge().Style
                .Font.SetBold().Font.SetFontSize(14)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            ws.Cell(3, 1).Value = "ФИО:";
            ws.Cell(3, 2).Value = fio;
            ws.Cell(4, 1).Value = "Номер группы:";
            ws.Cell(4, 2).Value = groupName;
            ws.Cell(5, 1).Value = "Количество пройденных тестов:";
            ws.Cell(5, 2).Value = rows.Count;
            ws.Range(3, 1, 5, 1).Style.Font.SetBold();

            int headerRow = 7;
            string[] headers =
            {
                "№ теста", "Дата сдачи", "Результат теста",
                "Максимальное количество баллов", "Процент верных ответов"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(headerRow, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.SetBold().Font.SetFontColor(XLColor.White);
                cell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(129, 166, 198));
                cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                cell.Style.Alignment.SetWrapText(true);
            }

            int r = headerRow + 1;
            foreach (var row in rows)
            {
                ws.Cell(r, 1).Value = row.TestNumber;
                ws.Cell(r, 2).Value = row.Date?.ToString("dd.MM.yyyy") ?? "—";
                ws.Cell(r, 3).Value = row.Score;
                ws.Cell(r, 4).Value = row.MaxScore;
                ws.Cell(r, 5).Value = row.Percent;
                ws.Cell(r, 5).Style.NumberFormat.Format = "0.00";

                ws.Cell(r, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Cell(r, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Cell(r, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Cell(r, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Cell(r, 5).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                r++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
        }

        private void BuildGroupExcel(string path, string groupName, List<StudentResult> groupResults)
        {
            var rows = BuildGroupRows(groupResults);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Отчёт по группе");

            ws.Cell(1, 1).Value = $"Отчёт по группе {groupName}";
            ws.Range(1, 1, 1, 5).Merge().Style
                .Font.SetBold().Font.SetFontSize(14)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            int headerRow = 3;
            string[] headers =
            {
                "№ теста", "Тема",
                "Сколько людей сдавали тест",
                "Сколько было попыток сдать тест",
                "Среднее количество баллов по всем попыткам"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(headerRow, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.SetBold().Font.SetFontColor(XLColor.White);
                cell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(129, 166, 198));
                cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                cell.Style.Alignment.SetWrapText(true);
            }

            int r = headerRow + 1;
            foreach (var row in rows)
            {
                ws.Cell(r, 1).Value = row.TestNumber;
                ws.Cell(r, 2).Value = row.Theme;
                ws.Cell(r, 3).Value = row.StudentsCount;
                ws.Cell(r, 4).Value = row.AttemptsCount;
                ws.Cell(r, 5).Value = row.AvgScore;
                ws.Cell(r, 5).Style.NumberFormat.Format = "0.00";

                ws.Cell(r, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Cell(r, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Cell(r, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Cell(r, 5).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                r++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
        }
    }
}