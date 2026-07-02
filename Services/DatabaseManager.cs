using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Access.Dao;
using TPDSSDataManager.Models;

namespace TPDSSDataManager.Services
{
    public class DatabaseManager
    {
        private readonly string[] _tables = { "ТП ДСС(4рівень)", "ТП ДСС_123" };

        // Метод анализа базы данных и построения дерева
        public async Task<ObservableCollection<TreeNode>> LoadSspTreeAsync(string path)
        {
            var rootList = new ObservableCollection<TreeNode>();

            await Task.Run(() =>
            {
                DBEngine engine = new DBEngine();
                Database db = engine.OpenDatabase(path);

                string oldName = "04. Управління координації";
                string newName = "04. Управління координації процесу збирання даних";
                foreach (var tableName in _tables)
                {
                    try { db.Execute($"UPDATE [{tableName}] SET [Назва ССП] = '{newName}' WHERE [Назва ССП] = '{oldName}'"); } catch { }
                }

                var dict = new Dictionary<string, HashSet<string>>();

                foreach (var tableName in _tables)
                {
                    try
                    {
                        Recordset rs = db.OpenRecordset($"SELECT DISTINCT [Назва ССП], [Назва СП] FROM [{tableName}] WHERE [Назва ССП] IS NOT NULL");
                        while (!rs.EOF)
                        {
                            string? ssp = rs.Fields["Назва ССП"].Value?.ToString()?.Trim();
                            string? sp = "";

                            try { sp = rs.Fields["Назва СП"].Value?.ToString()?.Trim(); } catch { }

                            if (!string.IsNullOrEmpty(ssp))
                            {
                                if (!dict.ContainsKey(ssp)) dict[ssp] = new HashSet<string>();
                                if (!string.IsNullOrEmpty(sp)) dict[ssp].Add(sp);
                            }
                            rs.MoveNext();
                        }
                        rs.Close();
                    }
                    catch { }
                }
                db.Close();

                foreach (var kvp in dict.OrderBy(x => x.Key))
                {
                    var parentNode = new TreeNode { Name = kvp.Key };

                    if (kvp.Value.Count > 0 && !(kvp.Value.Count == 1 && kvp.Value.First() == kvp.Key))
                    {
                        foreach (var sp in kvp.Value.OrderBy(x => x))
                        {
                            parentNode.Children.Add(new TreeNode { Name = sp, Parent = parentNode });
                        }
                    }
                    rootList.Add(parentNode);
                }
            });

            return rootList;
        }

        // Метод розбивки (Split)
        public async Task SplitDatabasesAsync(string sourcePath, string outputDir, List<(string SspName, string? SpName, string Prefix)> tasksToRun, IProgress<string> progress)
        {
            await Task.Run(() =>
            {
                var logger = new SimpleLogger(outputDir);
                DBEngine engine = new DBEngine();

                for (int i = 0; i < tasksToRun.Count; i++)
                {
                    var task = tasksToRun[i];
                    string num = string.IsNullOrEmpty(task.Prefix) ? $"upd_{i}" : task.Prefix;
                    string newPath = Path.Combine(outputDir, $"pl_srv_{num}.accdb");

                    try
                    {
                        File.Copy(sourcePath, newPath, true);
                        FilterSspInFile(engine, newPath, task.SspName, task.SpName, logger);
                        progress?.Report($"Створено: pl_srv_{num}.accdb");
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Split - {num}", ex.Message);
                        progress?.Report($"Помилка створення: pl_srv_{num}.accdb");
                    }
                }
            });
        }

        private void FilterSspInFile(DBEngine engine, string path, string sspName, string? spName, SimpleLogger logger)
        {
            Database db = engine.OpenDatabase(path);
            string cleanSsp = sspName.Replace("'", "''");

            foreach (var tableName in _tables)
            {
                try
                {
                    db.Execute($"DELETE FROM [{tableName}] WHERE [Назва ССП] <> '{cleanSsp}'");

                    if (!string.IsNullOrEmpty(spName))
                    {
                        string cleanSp = spName.Replace("'", "''");
                        db.Execute($"DELETE FROM [{tableName}] WHERE [Назва СП] <> '{cleanSp}' AND [Назва СП] IS NOT NULL");
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError($"FilterSspInFile ({tableName})", ex.Message);
                }
            }
            db.Close();
        }

        // Метод об'єднання (Merge)
        public async Task MergeDatabasesAsync(List<string> mergeFiles, string resultPath, IProgress<string> progress)
        {
            if (mergeFiles == null || mergeFiles.Count == 0) return;

            await Task.Run(() =>
            {
                string targetDir = Path.GetDirectoryName(resultPath) ?? string.Empty;
                var logger = new SimpleLogger(targetDir);
                DBEngine engine = new DBEngine();

                try
                {
                    // 1. Берем ПЕРВЫЙ файл как основу (сохраняем оригинальные ID и структуры)
                    File.Copy(mergeFiles[0], resultPath, true);
                }
                catch (Exception ex)
                {
                    logger.LogError("Merge - Копіювання базового файлу", ex.Message);
                    progress?.Report("Помилка створення базового файлу!");
                    return;
                }

                // 2. Начинаем цикл со ВТОРОГО файла (индекс 1)
                for (int i = 1; i < mergeFiles.Count; i++)
                {
                    string sourceFile = mergeFiles[i];
                    string fileName = Path.GetFileName(sourceFile);
                    progress?.Report($"Обробка: {fileName}");

                    Database? dbDst = null;
                    try
                    {
                        dbDst = engine.OpenDatabase(resultPath);

                        foreach (var tableName in _tables)
                        {
                            try
                            {
                                Recordset rsDstDef = dbDst.OpenRecordset($"SELECT * FROM [{tableName}] WHERE 1=0");

                                List<string> normalFields = new List<string>();
                                List<string> complexFields = new List<string>();

                                for (int f = 0; f < rsDstDef.Fields.Count; f++)
                                {
                                    Field fld = rsDstDef.Fields[f];
                                    if ((short)fld.Type >= 101)
                                        complexFields.Add(fld.Name);
                                    else
                                        normalFields.Add(fld.Name);
                                }
                                rsDstDef.Close();

                                // Шаг А: Обход блокировки IN через создание временной привязки (Linked Table)
                                if (normalFields.Count > 0)
                                {
                                    string linkedTableName = $"Link_{Guid.NewGuid().ToString("N").Substring(0, 6)}";

                                    // Программно линкуем таблицу из внешнего файла
                                    TableDef tdf = dbDst.CreateTableDef(linkedTableName);
                                    tdf.Connect = $";DATABASE={sourceFile}";
                                    tdf.SourceTableName = tableName;
                                    dbDst.TableDefs.Append(tdf);

                                    try
                                    {
                                        string fieldsList = string.Join(", ", normalFields.Select(f => $"[{f}]"));
                                        // Теперь IN не нужен, Access думает, что таблица локальная
                                        string sqlInsert = $"INSERT INTO [{tableName}] ({fieldsList}) SELECT {fieldsList} FROM [{linkedTableName}]";
                                        dbDst.Execute(sqlInsert, 128);
                                    }
                                    finally
                                    {
                                        // Обязательно удаляем линк после вставки, даже если была ошибка
                                        dbDst.TableDefs.Delete(linkedTableName);
                                    }
                                }

                                // Шаг Б: Дописываем чекбоксы/MVF
                                if (complexFields.Count > 0)
                                {
                                    Database dbSrc = engine.OpenDatabase(sourceFile);
                                    Recordset rsSrc = dbSrc.OpenRecordset($"SELECT * FROM [{tableName}]");

                                    string pkName = normalFields.FirstOrDefault(f => f.Equals("Ідентифікатор", StringComparison.OrdinalIgnoreCase)) ?? normalFields[0];

                                    while (!rsSrc.EOF)
                                    {
                                        object pkValue = rsSrc.Fields[pkName].Value;
                                        if (pkValue != null)
                                        {
                                            string pkCondition = (pkValue is string) ? $"'{pkValue.ToString().Replace("'", "''")}'" : pkValue.ToString();
                                            Recordset rsDst = dbDst.OpenRecordset($"SELECT * FROM [{tableName}] WHERE [{pkName}] = {pkCondition}");

                                            if (!rsDst.EOF)
                                            {
                                                rsDst.Edit();
                                                foreach (string cField in complexFields)
                                                {
                                                    Recordset2 childSrc = (Recordset2)rsSrc.Fields[cField].Value;
                                                    Recordset2 childDst = (Recordset2)rsDst.Fields[cField].Value;

                                                    if (childSrc != null && !childSrc.EOF)
                                                    {
                                                        childSrc.MoveFirst();
                                                        while (!childSrc.EOF)
                                                        {
                                                            childDst.AddNew();
                                                            for (int c = 0; c < childSrc.Fields.Count; c++)
                                                            {
                                                                try { childDst.Fields[c].Value = childSrc.Fields[c].Value; }
                                                                catch (Exception fieldEx) { logger.LogError($"Merge MVF Field ({tableName} -> {cField})", fieldEx.Message); }
                                                            }
                                                            childDst.Update();
                                                            childSrc.MoveNext();
                                                        }
                                                    }
                                                }
                                                rsDst.Update();
                                            }
                                            rsDst.Close();
                                        }
                                        rsSrc.MoveNext();
                                    }
                                    rsSrc.Close();
                                    dbSrc.Close();
                                }
                            }
                            catch (Exception ex)
                            {
                                logger.LogError($"Merge Table ({tableName}) in {fileName}", ex.Message);
                                progress?.Report($"Помилка у таблиці {tableName}. Див. лог.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Merge File Access ({fileName})", ex.Message);
                    }
                    finally
                    {
                        dbDst?.Close();
                    }
                }
                progress?.Report("Склейка завершена!");
            });
        }
    }
}