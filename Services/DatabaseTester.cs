using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Access.Dao;

namespace TPDSSDataManager.Services
{
    public class DatabaseTester
    {
        private readonly string[] _tables = { "ТП ДСС(4рівень)", "ТП ДСС_123" };

        public async Task<string> RunTestAsync(List<string> sourceFiles, string mergedFile, IProgress<string> progress)
        {
            return await Task.Run(() =>
            {
                StringBuilder report = new StringBuilder();
                DBEngine engine = new DBEngine();
                Database? dbMerged = null;

                try
                {
                    progress?.Report("Відкриття фінальної бази для перевірки...");
                    dbMerged = engine.OpenDatabase(mergedFile);

                    foreach (var tableName in _tables)
                    {
                        report.AppendLine($"=== Аналіз таблиці: [{tableName}] ===");
                        progress?.Report($"Аналіз таблиці: {tableName}");

                        string pkName = "Ідентифікатор";
                        List<string> normalFields = new List<string>();

                        // 1. Отримуємо список полів ОДИН раз
                        try
                        {
                            Recordset rsDef = dbMerged.OpenRecordset($"SELECT * FROM [{tableName}] WHERE 1=0");
                            for (int f = 0; f < rsDef.Fields.Count; f++)
                            {
                                Field fld = rsDef.Fields[f];
                                if ((short)fld.Type < 101)
                                {
                                    normalFields.Add(fld.Name);
                                    if (fld.Name.Equals("Ідентифікатор", StringComparison.OrdinalIgnoreCase)) pkName = fld.Name;
                                }
                            }
                            rsDef.Close();
                            Marshal.ReleaseComObject(rsDef);
                        }
                        catch { continue; }

                        HashSet<string> mergedRowsContent = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        // 2. Читаємо фінальну базу (ОПТИМІЗОВАНО)
                        try
                        {
                            Recordset rsMerged = dbMerged.OpenRecordset($"SELECT * FROM [{tableName}]");

                            // Кешуємо об'єкти полів, щоб не шукати їх по імені кожен рядок
                            List<Field> cachedFields = new List<Field>();
                            foreach (var fldName in normalFields)
                            {
                                if (!fldName.Equals(pkName, StringComparison.OrdinalIgnoreCase))
                                    cachedFields.Add(rsMerged.Fields[fldName]);
                            }

                            while (!rsMerged.EOF)
                            {
                                mergedRowsContent.Add(BuildFastKey(cachedFields));
                                rsMerged.MoveNext();
                            }
                            rsMerged.Close();
                            Marshal.ReleaseComObject(rsMerged);
                        }
                        catch (Exception ex)
                        {
                            report.AppendLine($"Помилка читання фінальної бази: {ex.Message}\n");
                            continue;
                        }

                        int totalSourceRows = 0;
                        List<string> missingDetails = new List<string>();

                        // 3. Читаємо файли співробітників (ОПТИМІЗОВАНО)
                        for (int i = 0; i < sourceFiles.Count; i++)
                        {
                            string srcFile = sourceFiles[i];
                            progress?.Report($"Перевірка {i + 1}/{sourceFiles.Count}: {Path.GetFileName(srcFile)}");
                            Database? dbSrc = null;
                            Recordset? rsSrc = null;

                            try
                            {
                                dbSrc = engine.OpenDatabase(srcFile);
                                rsSrc = dbSrc.OpenRecordset($"SELECT * FROM [{tableName}]");

                                // Кешуємо об'єкти полів для поточного файлу
                                List<Field> cachedSrcFields = new List<Field>();
                                Field? pkField = null;
                                Field? sspField = null;

                                foreach (var fldName in normalFields)
                                {
                                    if (fldName.Equals(pkName, StringComparison.OrdinalIgnoreCase))
                                        pkField = rsSrc.Fields[fldName];
                                    else
                                        cachedSrcFields.Add(rsSrc.Fields[fldName]);

                                    if (fldName.Equals("Назва ССП", StringComparison.OrdinalIgnoreCase))
                                        sspField = rsSrc.Fields[fldName];
                                }

                                while (!rsSrc.EOF)
                                {
                                    totalSourceRows++;
                                    string srcKey = BuildFastKey(cachedSrcFields);

                                    if (!mergedRowsContent.Contains(srcKey))
                                    {
                                        string idVal = pkField?.Value?.ToString() ?? "N/A";
                                        string ssp = sspField?.Value?.ToString() ?? "-";
                                        missingDetails.Add($"[ВТРАЧЕНО] Ориг. ID: {idVal,-5} | ССП: {ssp} (Файл: {Path.GetFileName(srcFile)})");
                                    }
                                    rsSrc.MoveNext();
                                }
                            }
                            catch (Exception ex)
                            {
                                report.AppendLine($"Помилка у файлі {Path.GetFileName(srcFile)}: {ex.Message}");
                            }
                            finally
                            {
                                if (rsSrc != null) { try { rsSrc.Close(); Marshal.ReleaseComObject(rsSrc); } catch { } }
                                if (dbSrc != null) { try { dbSrc.Close(); Marshal.ReleaseComObject(dbSrc); } catch { } }
                            }
                        }

                        report.AppendLine($"Всього рядків у фінальній базі: {mergedRowsContent.Count}");
                        report.AppendLine($"Всього рядків у шматках (сума): {totalSourceRows}");
                        report.AppendLine($"❌ Знайдено втрачених рядків контенту: {missingDetails.Count}");

                        if (missingDetails.Count > 0)
                        {
                            report.AppendLine("\n--- ДЕТАЛІ ВТРАЧЕНИХ РЯДКІВ ---");
                            for (int j = 0; j < Math.Min(missingDetails.Count, 300); j++)
                                report.AppendLine(missingDetails[j]);
                            if (missingDetails.Count > 300) report.AppendLine("... та інші (виведено перші 300).");
                        }
                        report.AppendLine();
                    }

                    progress?.Report("Тестування завершено!");
                    return report.ToString();
                }
                catch (Exception ex) { return $"КРИТИЧНА ПОМИЛКА: {ex.Message}"; }
                finally
                {
                    if (dbMerged != null) { try { dbMerged.Close(); Marshal.ReleaseComObject(dbMerged); } catch { } }
                    if (engine != null) { try { Marshal.ReleaseComObject(engine); } catch { } }
                }
            });
        }

        // Швидкий метод збірки рядка без пошуку колонок по назві
        private static string BuildFastKey(List<Field> cachedFields)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var fld in cachedFields)
            {
                sb.Append(fld.Value?.ToString()?.Trim() ?? "").Append("§");
            }
            return sb.ToString();
        }
    }
}