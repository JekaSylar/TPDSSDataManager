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
                    // Добавляем без Dispatcher, так как мы вернем коллекцию в UI-поток после await
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
                DBEngine engine = new DBEngine();

                for (int i = 0; i < tasksToRun.Count; i++)
                {
                    var task = tasksToRun[i];
                    string num = string.IsNullOrEmpty(task.Prefix) ? $"upd_{i}" : task.Prefix;
                    string newPath = Path.Combine(outputDir, $"pl_srv_{num}.accdb");

                    File.Copy(sourcePath, newPath, true);
                    FilterSspInFile(engine, newPath, task.SspName, task.SpName);

                    // Сообщаем UI потоку об обновлении статуса
                    progress?.Report($"Створено: pl_srv_{num}.accdb");
                }
            });
        }

        private void FilterSspInFile(DBEngine engine, string path, string sspName, string? spName)
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
                        try { db.Execute($"DELETE FROM [{tableName}] WHERE [Назва СП] <> '{cleanSp}' AND [Назва СП] IS NOT NULL"); } catch { }
                    }
                }
                catch { }
            }
            db.Close();
        }

        // Метод об'єднання (Merge)
        public async Task MergeDatabasesAsync(List<string> mergeFiles, string resultPath, IProgress<string> progress)
        {
            await Task.Run(() =>
            {
                DBEngine engine = new DBEngine();
                File.Copy(mergeFiles[0], resultPath, true);
                Database dbRes = engine.OpenDatabase(resultPath);

                foreach (var t in _tables)
                {
                    try { dbRes.Execute($"DELETE * FROM [{t}]"); } catch { }
                }
                dbRes.Close();

                foreach (var sourceFile in mergeFiles)
                {
                    progress?.Report($"Обробка: {Path.GetFileName(sourceFile)}");

                    Database dbSrc = engine.OpenDatabase(sourceFile);
                    Database dbDst = engine.OpenDatabase(resultPath);

                    foreach (var tableName in _tables)
                    {
                        try
                        {
                            Recordset rsSrc = dbSrc.OpenRecordset($"SELECT * FROM [{tableName}]");
                            Recordset rsDst = dbDst.OpenRecordset($"SELECT * FROM [{tableName}]");

                            while (!rsSrc.EOF)
                            {
                                rsDst.AddNew();
                                for (int i = 1; i < rsSrc.Fields.Count; i++)
                                {
                                    Field fld = rsSrc.Fields[i];
                                    try
                                    {
                                        if ((short)fld.Type >= 101)
                                        {
                                            Recordset2 childSrc = (Recordset2)fld.Value;
                                            Recordset2 childDst = (Recordset2)rsDst.Fields[fld.Name].Value;

                                            if (childSrc != null)
                                            {
                                                while (!childSrc.EOF)
                                                {
                                                    childDst.AddNew();
                                                    childDst.Fields[0].Value = childSrc.Fields[0].Value;
                                                    childDst.Update();
                                                    childSrc.MoveNext();
                                                }
                                            }
                                        }
                                        else
                                        {
                                            rsDst.Fields[fld.Name].Value = fld.Value;
                                        }
                                    }
                                    catch { }
                                }
                                rsDst.Update();
                                rsSrc.MoveNext();
                            }
                            rsSrc.Close();
                            rsDst.Close();
                        }
                        catch { }
                    }
                    dbSrc.Close();
                    dbDst.Close();
                }
            });
        }
    }
}