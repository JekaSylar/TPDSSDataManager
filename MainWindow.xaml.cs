using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using TPDSSDataManager.Models;
using TPDSSDataManager.Services;
using static System.Net.WebRequestMethods;

namespace TPDSSDataManager
{
    public partial class MainWindow : Window
    {
        private string _sourceSplitPath = string.Empty;
        private List<string> _mergeFiles = new List<string>();
        private ObservableCollection<TreeNode> _treeNodes = new ObservableCollection<TreeNode>();

        // Змінні для тестування
        private string _testMergedPath = string.Empty;
        private List<string> _testSourceFiles = new List<string>();
        private readonly DatabaseTester _dbTester = new DatabaseTester();

        // Змінна для відміни склейки
        private CancellationTokenSource? _mergeCancellationTokenSource;

        private readonly DatabaseManager _dbManager;
    

        private readonly string CURRENT_VERSION = "1.5.1";

        public MainWindow()
        {
            

            InitializeComponent();
            _dbManager = new DatabaseManager();

          
        }

        // --- МЕНЮ ТА НАВІГАЦІЯ ---
        private void MenuSplit_Click(object sender, RoutedEventArgs e) => ShowPanel(GridSplit, "Режим: Розбивка");
        private void MenuMerge_Click(object sender, RoutedEventArgs e) => ShowPanel(GridMerge, "Режим: Об'єднання");
        private void MenuTest_Click(object sender, RoutedEventArgs e) => ShowPanel(GridTest, "Режим: Тестування склейки");
        private void MenuExit_Click(object sender, RoutedEventArgs e) => Application.Current.Shutdown();

        private void ShowPanel(UIElement panel, string status)
        {
            StartScreen.Visibility = Visibility.Collapsed;
            GridSplit.Visibility = Visibility.Collapsed;
            GridMerge.Visibility = Visibility.Collapsed;
            GridTest.Visibility = Visibility.Collapsed;

            panel.Visibility = Visibility.Visible;
            TxtStatus.Text = status;
        }

        // --- ЛОГІКА РОЗБИВКИ (Split) ---
        private async void BtnSelectSource_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "Access Files (*.accdb)|*.accdb" };
            if (dlg.ShowDialog() == true)
            {
                _sourceSplitPath = dlg.FileName;
                TxtSourcePath.Text = Path.GetFileName(_sourceSplitPath);
                BtnRunSplit.IsEnabled = false;

                SetLoading(true, "Аналіз та виправлення бази даних...");

                _treeNodes = await _dbManager.LoadSspTreeAsync(_sourceSplitPath);

                TreeSSP.ItemsSource = _treeNodes;
                TreePanel.Visibility = Visibility.Visible;
                BtnRunSplit.IsEnabled = _treeNodes.Count > 0;

                SetLoading(false, "Готово");
            }
        }

        private void ChkSelectAll_Click(object sender, RoutedEventArgs e)
        {
            bool state = ChkSelectAll.IsChecked == true;
            foreach (var node in _treeNodes)
            {
                node.IsChecked = state;
            }
        }

        private async void BtnRunSplit_Click(object sender, RoutedEventArgs e)
        {
            var tasksToRun = new List<(string SspName, string? SpName, string Prefix)>();

            foreach (var parent in _treeNodes)
            {
                if (parent.IsChecked == true)
                {
                    string prefix = Regex.Match(parent.Name, @"^\d+(?:\.\d+)?").Value;
                    tasksToRun.Add((parent.Name, null, prefix));
                }
                else if (parent.IsChecked == null)
                {
                    foreach (var child in parent.Children)
                    {
                        if (child.IsChecked == true)
                        {
                            string prefix = Regex.Match(child.Name, @"^\d+(?:\.\d+)?").Value;
                            tasksToRun.Add((parent.Name, child.Name, prefix));
                        }
                    }
                }
            }

            if (tasksToRun.Count == 0)
            {
                MessageBox.Show("Оберіть хоча б один підрозділ для розбивки!", "Увага", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var folderDlg = new Microsoft.Win32.OpenFileDialog { CheckFileExists = false, FileName = "Вибір папки", Title = "Куди зберігати?" };
            if (folderDlg.ShowDialog() != true) return;

            string outputDir = Path.GetDirectoryName(folderDlg.FileName) ?? "";

            SetLoading(true, "Виконується розбивка...");
            try
            {
                var progress = new Progress<string>(status => TxtSubStatus.Text = status);
                await _dbManager.SplitDatabasesAsync(_sourceSplitPath, outputDir, tasksToRun, progress);

                TxtStatus.Text = "✅ Розбивку завершено!";
                MessageBox.Show($"Успішно створено файлів: {tasksToRun.Count}", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка при розбивці: " + ex.Message);
            }
            finally
            {
                SetLoading(false, "Готово");
            }
        }

        // --- ЛОГІКА ОБ'ЄДНАННЯ (Merge) ---
        private void BtnAddMergeFiles_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Multiselect = true, Filter = "Access|*.accdb" };
            if (dlg.ShowDialog() == true) AddFilesToMergeList(dlg.FileNames);
        }

        private void BtnClearMergeFiles_Click(object sender, RoutedEventArgs e)
        {
            _mergeFiles.Clear();
            ListMergeFiles.Items.Clear();
            BtnRunMerge.IsEnabled = false;
        }

        // Видалення виділених файлів зі списку злиття (по кнопці Del або контекстному меню)
        private void ListMergeFiles_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete) RemoveSelectedMergeFiles();
        }

        private void RemoveSelectedMergeFiles_Click(object sender, RoutedEventArgs e) => RemoveSelectedMergeFiles();

        private void RemoveSelectedMergeFiles()
        {
            int index = ListMergeFiles.SelectedIndex;
            if (index >= 0)
            {
                _mergeFiles.RemoveAt(index);
                ListMergeFiles.Items.RemoveAt(index);
                BtnRunMerge.IsEnabled = _mergeFiles.Count > 0;
            }
        }

        private void ListMergeFiles_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }

        private void ListMergeFiles_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                AddFilesToMergeList(files);
            }
        }

        private void AddFilesToMergeList(string[] files)
        {
            bool hasNewFiles = false;
            foreach (var file in files)
            {
                if (Path.GetExtension(file).Equals(".accdb", StringComparison.OrdinalIgnoreCase) && !_mergeFiles.Contains(file))
                {
                    _mergeFiles.Add(file);
                    ListMergeFiles.Items.Add(Path.GetFileName(file));
                    hasNewFiles = true;
                }
            }
            if (hasNewFiles) BtnRunMerge.IsEnabled = true;
        }

        private async void BtnRunMerge_Click(object sender, RoutedEventArgs e)
        {
            var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "Access|*.accdb", FileName = "result.accdb" };
            if (sfd.ShowDialog() != true) return;

            string resultPath = sfd.FileName;
            _mergeCancellationTokenSource = new CancellationTokenSource();

            SetLoading(true, "Об'єднання баз...");
            BtnCancelMerge.IsEnabled = true;
            BtnRunMerge.IsEnabled = false;
            BtnClearMergeFiles.IsEnabled = false;

            try
            {
                var progress = new Progress<string>(status => TxtSubStatus.Text = status);
                await _dbManager.MergeDatabasesAsync(_mergeFiles, resultPath, progress, _mergeCancellationTokenSource.Token);

                TxtStatus.Text = "✅ Об'єднання завершено!";
                MessageBox.Show("Об'єднання файлів успішно завершено!", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (OperationCanceledException)
            {
                TxtStatus.Text = "⚠️ Процес скасовано користувачем.";
                MessageBox.Show("Склейку перервано. Створений файл може бути неповним.", "Скасовано", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка при об'єднанні: " + ex.Message);
            }
            finally
            {
                BtnCancelMerge.IsEnabled = false;
                BtnClearMergeFiles.IsEnabled = true;
                _mergeCancellationTokenSource?.Dispose();
                _mergeCancellationTokenSource = null;
                SetLoading(false, TxtStatus.Text);
            }
        }

        private void BtnCancelMerge_Click(object sender, RoutedEventArgs e)
        {
            if (_mergeCancellationTokenSource != null && !_mergeCancellationTokenSource.IsCancellationRequested)
            {
                _mergeCancellationTokenSource.Cancel();
                BtnCancelMerge.IsEnabled = false;
                TxtStatus.Text = "Зупинка процесу, зачекайте...";
            }
        }

        // --- ЛОГІКА ТЕСТУВАННЯ ---
        private void BtnSelectTestMerged_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "Access Files (*.accdb)|*.accdb" };
            if (dlg.ShowDialog() == true)
            {
                _testMergedPath = dlg.FileName;
                TxtTestMergedPath.Text = Path.GetFileName(_testMergedPath);
                CheckTestReady();
            }
        }

        private void BtnAddTestFiles_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Multiselect = true, Filter = "Access|*.accdb" };
            if (dlg.ShowDialog() == true) AddTestFiles(dlg.FileNames);
        }

        private void BtnClearTestFiles_Click(object sender, RoutedEventArgs e)
        {
            _testSourceFiles.Clear();
            ListTestFiles.Items.Clear();
            CheckTestReady();
        }

        // Видалення виділених файлів зі списку тестування
        private void ListTestFiles_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete) RemoveSelectedTestFiles();
        }

        private void RemoveSelectedTestFiles_Click(object sender, RoutedEventArgs e) => RemoveSelectedTestFiles();

        private void RemoveSelectedTestFiles()
        {
            int index = ListTestFiles.SelectedIndex;
            if (index >= 0)
            {
                _testSourceFiles.RemoveAt(index);
                ListTestFiles.Items.RemoveAt(index);
                CheckTestReady();
            }
        }

        private void ListTestFiles_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effects = DragDropEffects.Copy;
            else e.Effects = DragDropEffects.None;
        }

        private void ListTestFiles_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                AddTestFiles((string[])e.Data.GetData(DataFormats.FileDrop));
            }
        }

        private void AddTestFiles(string[] files)
        {
            foreach (var file in files)
            {
                if (Path.GetExtension(file).Equals(".accdb", StringComparison.OrdinalIgnoreCase) && !_testSourceFiles.Contains(file))
                {
                    _testSourceFiles.Add(file);
                    ListTestFiles.Items.Add(Path.GetFileName(file));
                }
            }
            CheckTestReady();
        }

        private void CheckTestReady()
        {
            BtnRunTest.IsEnabled = !string.IsNullOrEmpty(_testMergedPath) && _testSourceFiles.Count > 0;
        }

        private async void BtnRunTest_Click(object sender, RoutedEventArgs e)
        {
            SetLoading(true, "Йде глибоке порівняння баз...");
            TxtTestReport.Text = "Аналізуємо... Це може зайняти хвилину...";

            try
            {
                var progress = new Progress<string>(status => TxtSubStatus.Text = status);
                string report = await _dbTester.RunTestAsync(_testSourceFiles, _testMergedPath, progress);
                TxtTestReport.Text = report;
                TxtStatus.Text = "✅ Тестування завершено!";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка під час тестування: " + ex.Message);
            }
            finally
            {
                SetLoading(false, "Готово");
            }
        }

        private void SetLoading(bool isLoading, string status)
        {
            MainProgress.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
            if (TxtStatus.Text != "⚠️ Процес скасовано користувачем." || isLoading)
            {
                TxtStatus.Text = status;
            }

            BtnRunSplit.IsEnabled = !isLoading && _treeNodes.Count > 0;
            BtnRunMerge.IsEnabled = !isLoading && _mergeFiles.Count > 0;
        }

        private void MenuHelp_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show($"ТП ДСС: Автоматизація ГУС у Донецькій області\nВерсія програми: v {CURRENT_VERSION}", "Довідка", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}