using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using StudentFileWorkTask.data;

namespace StudentFileWorkTask.presentation
{
    public partial class ColumnMappingWindow : Window
    {
        private List<string> _headers;

        public MappingTemplate ResultTemplate { get; private set; }

        public ColumnMappingWindow(List<string> headers)
        {
            InitializeComponent();
            _headers = headers;

            cmbStudent.ItemsSource = _headers;
            cmbGroup.ItemsSource = _headers;
            cmbDate.ItemsSource = _headers;
        }

        private void AddQuestion_Click(object sender, RoutedEventArgs e)
        {
            var pair = CreateQuestionScorePair();
            questionsPanel.Children.Add(pair);
        }

        private StackPanel CreateQuestionScorePair(string selectedQuestion = null, string selectedScore = null)
        {
            var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 5) };

            var cmbQuestion = new ComboBox
            {
                Width = 200,
                Margin = new Thickness(0, 0, 10, 0),
                ItemsSource = _headers,
                Tag = "Question"
            };
            if (selectedQuestion != null) cmbQuestion.SelectedItem = selectedQuestion;

            var cmbScore = new ComboBox
            {
                Width = 200,
                ItemsSource = _headers,
                Tag = "Score"
            };
            if (selectedScore != null) cmbScore.SelectedItem = selectedScore;

            var removeBtn = new Button
            {
                Content = "✕",
                Width = 30,
                Height = 25,
                Margin = new Thickness(5, 0, 0, 0),
                Background = System.Windows.Media.Brushes.LightGray
            };
            removeBtn.Click += (s, args) => questionsPanel.Children.Remove(panel);

            panel.Children.Add(new TextBlock { Text = "Вопрос:", Width = 50, VerticalAlignment = VerticalAlignment.Center });
            panel.Children.Add(cmbQuestion);
            panel.Children.Add(new TextBlock { Text = "Балл:", Width = 40, VerticalAlignment = VerticalAlignment.Center });
            panel.Children.Add(cmbScore);
            panel.Children.Add(removeBtn);

            return panel;
        }

        private void AutoMapping_Click(object sender, RoutedEventArgs e)
        {
            questionsPanel.Children.Clear();

            var excludeKeywords = new[] {
                "id", "время", "дата", "групп", "фамилия", "имя", "отчество", "фио",
                "балл", "набрано", "всего", "результат", "определение", "машин"
            };

            var questionColumns = new List<string>();
            var scoreColumns = new List<string>();

            foreach (var col in _headers)
            {
                if (string.IsNullOrEmpty(col)) continue;

                var lower = col.ToLower();
                bool isExcluded = false;

                foreach (var kw in excludeKeywords)
                {
                    if (lower.Contains(kw))
                    {
                        isExcluded = true;
                        break;
                    }
                }

                if (!isExcluded)
                {
                    questionColumns.Add(col);
                }

                if (lower.Contains("балл") || lower.Contains("score"))
                {
                    scoreColumns.Add(col);
                }
            }

            for (int i = 0; i < questionColumns.Count; i++)
            {
                string score = null;
                if (i < scoreColumns.Count)
                {
                    score = scoreColumns[i];
                }
                questionsPanel.Children.Add(CreateQuestionScorePair(questionColumns[i], score));
            }

            foreach (var col in _headers)
            {
                var lower = col.ToLower();
                if (lower.Contains("фамилия") || lower.Contains("фио") || lower.Contains("имя") || lower.Contains("отчество"))
                    cmbStudent.SelectedItem = col;
                if (lower.Contains("групп") || lower.Contains("номер"))
                    cmbGroup.SelectedItem = col;
                if (lower.Contains("время") || lower.Contains("дата") || lower.Contains("создания"))
                    cmbDate.SelectedItem = col;
            }

            if (questionColumns.Count == 0)
            {
                for (int i = 0; i < 5; i++)
                {
                    questionsPanel.Children.Add(CreateQuestionScorePair(null, null));
                }
                MessageBox.Show("Не удалось определить вопросы. Добавьте их вручную.",
                    "Автоопределение", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private bool ValidateMapping()
        {
            if (string.IsNullOrEmpty(cmbStudent.SelectedItem as string))
            {
                MessageBox.Show("Выберите столбец для студента!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (string.IsNullOrEmpty(cmbGroup.SelectedItem as string))
            {
                MessageBox.Show("Выберите столбец для группы!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (questionsPanel.Children.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы одну пару Вопрос-Балл!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            foreach (StackPanel panel in questionsPanel.Children)
            {
                var cmbQuestion = panel.Children.OfType<ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "Question");
                var cmbScore = panel.Children.OfType<ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "Score");

                if (cmbQuestion?.SelectedItem == null)
                {
                    MessageBox.Show("Не выбран столбец для вопроса!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }

                if (cmbScore?.SelectedItem == null)
                {
                    MessageBox.Show("Не выбран столбец для баллов!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }

            return true;
        }

        private void LoadTemplate_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog { Filter = "JSON|*.json" };
            if (dialog.ShowDialog() == true)
            {
                var json = File.ReadAllText(dialog.FileName);
                var template = JsonSerializer.Deserialize<MappingTemplate>(json);
                if (template != null)
                {
                    cmbStudent.SelectedItem = template.StudentColumn;
                    cmbGroup.SelectedItem = template.GroupColumn;
                    cmbDate.SelectedItem = template.DateColumn;

                    questionsPanel.Children.Clear();
                    for (int i = 0; i < template.QuestionColumns.Count; i++)
                    {
                        var score = i < template.ScoreColumns.Count ? template.ScoreColumns[i] : null;
                        questionsPanel.Children.Add(CreateQuestionScorePair(template.QuestionColumns[i], score));
                    }
                }
            }
        }

        private void SaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateMapping()) return;

            var dialog = new SaveFileDialog { Filter = "JSON|*.json", FileName = "mapping.json" };
            if (dialog.ShowDialog() == true)
            {
                var template = new MappingTemplate
                {
                    StudentColumn = cmbStudent.SelectedItem as string,
                    GroupColumn = cmbGroup.SelectedItem as string,
                    DateColumn = cmbDate.SelectedItem as string,
                    QuestionColumns = new List<string>(),
                    ScoreColumns = new List<string>()
                };

                foreach (StackPanel panel in questionsPanel.Children)
                {
                    var cmbQuestion = panel.Children.OfType<ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "Question");
                    var cmbScore = panel.Children.OfType<ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "Score");

                    if (cmbQuestion?.SelectedItem != null)
                        template.QuestionColumns.Add(cmbQuestion.SelectedItem.ToString());
                    if (cmbScore?.SelectedItem != null)
                        template.ScoreColumns.Add(cmbScore.SelectedItem.ToString());
                }

                var json = JsonSerializer.Serialize(template);
                File.WriteAllText(dialog.FileName, json);
                MessageBox.Show($"Шаблон успешно сохранён!");
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateMapping()) return;

            ResultTemplate = new MappingTemplate
            {
                StudentColumn = cmbStudent.SelectedItem as string,
                GroupColumn = cmbGroup.SelectedItem as string,
                DateColumn = cmbDate.SelectedItem as string,
                QuestionColumns = new List<string>(),
                ScoreColumns = new List<string>()
            };

            foreach (StackPanel panel in questionsPanel.Children)
            {
                var cmbQuestion = panel.Children.OfType<ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "Question");
                var cmbScore = panel.Children.OfType<ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "Score");

                if (cmbQuestion?.SelectedItem != null)
                    ResultTemplate.QuestionColumns.Add(cmbQuestion.SelectedItem.ToString());
                if (cmbScore?.SelectedItem != null)
                    ResultTemplate.ScoreColumns.Add(cmbScore.SelectedItem.ToString());
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}