using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using Microsoft.Win32;
using TPDSSDataManager.Models;
using TPDSSDataManager.Services;

namespace TPDSSDataManager
{
    public partial class MainWindow : Window
    {
        private string _sourceSplitPath = string.Empty;
        private List<string> _mergeFiles = new List<string>();
        private ObservableCollection<TreeNode> _treeNodes = new ObservableCollection<TreeNode>();

        private readonly DatabaseManager _dbManager;
        private readonly string CURRENT_VERSION = "1.4.0";

        public MainWindow()
        {
            InitializeComponent();
            _dbManager = new DatabaseManager(); // Инициализируем наш сервис работы с БД
        }

        // --- МЕНЮ ТА НАВІГАЦІЯ ---
        private void MenuSplit_Click(object sender, RoutedEventArgs e) => ShowPanel(GridSplit, "Режим: Розбивка");
        private void MenuMerge_Click(object sender, RoutedEventArgs e) => ShowPanel(GridMerge, "Режим: Об'єднання");
        private void MenuExit_Click(object sender, RoutedEventArgs e) => Application.Current.Shutdown();

        private void ShowPanel(UIElement panel, string status)
        {
            StartScreen.Visibility = Visibility.Collapsed;
            GridSplit.Visibility = Visibility.Collapsed;
            GridMerge.Visibility = Visibility.Collapsed;
            panel.Visibility = Visibility.Visible;
            TxtStatus.Text = status;
        }

        // --- ЛОГІКА РОЗБИВКИ (Split) ---
        private async void BtnSelectSource_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Access Files (*.accdb)|*.accdb" };
            if (dlg.ShowDialog() == true)
            {
                _sourceSplitPath = dlg.FileName;
                TxtSourcePath.Text = Path.GetFileName(_sourceSplitPath);
                BtnRunSplit.IsEnabled = false;

                SetLoading(true, "Аналіз та виправлення бази даних...");

                // Получаем дерево из сервиса
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

            var folderDlg = new OpenFileDialog { CheckFileExists = false, FileName = "Вибір папки", Title = "Куди зберігати?" };
            if (folderDlg.ShowDialog() != true) return;

            string outputDir = Path.GetDirectoryName(folderDlg.FileName) ?? "";

            SetLoading(true, "Виконується розбивка...");
            try
            {
                // Настраиваем объект для получения прогресса из фонового потока
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
            var dlg = new OpenFileDialog { Multiselect = true, Filter = "Access|*.accdb" };
            if (dlg.ShowDialog() == true)
            {
                AddFilesToMergeList(dlg.FileNames);
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

            if (hasNewFiles)
                BtnRunMerge.IsEnabled = true;
        }

        private async void BtnRunMerge_Click(object sender, RoutedEventArgs e)
        {
            var sfd = new SaveFileDialog { Filter = "Access|*.accdb", FileName = "result.accdb" };
            if (sfd.ShowDialog() != true) return;

            string resultPath = sfd.FileName;
            SetLoading(true, "Об'єднання баз (DAO)...");

            try
            {
                var progress = new Progress<string>(status => TxtSubStatus.Text = status);

                await _dbManager.MergeDatabasesAsync(_mergeFiles, resultPath, progress);

                TxtStatus.Text = "✅ Об'єднання завершено!";
                MessageBox.Show("Об'єднання файлів успішно завершено!", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка при об'єднанні: " + ex.Message);
            }
            finally
            {
                SetLoading(false, "Готово");
            }
        }

        private void SetLoading(bool isLoading, string status)
        {
            MainProgress.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
            TxtStatus.Text = status;

            BtnRunSplit.IsEnabled = !isLoading && _treeNodes.Count > 0;
            BtnRunMerge.IsEnabled = !isLoading && _mergeFiles.Count > 0;
        }

        private void MenuHelp_Click(object sender, RoutedEventArgs e) 
        { 
            MessageBox.Show($"ТП ДСС: Автоматизація ГУС у Донецькій області\nВерсія програми: v {CURRENT_VERSION}", "Довідка", MessageBoxButton.OK, MessageBoxImage.Information); 
        }
    
}
}