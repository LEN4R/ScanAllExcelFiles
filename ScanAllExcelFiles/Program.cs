using ClosedXML.Excel;
using ExcelDataReader;
using ExcelDataReader.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        // Нужно для .xlsb/.xls
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        Console.WriteLine("=== Сканирование Excel файлов (*.xlsx, *.xlsm, *.xlsb) ===\n");

        // 1. Папка
        string rootFolder = "";
        while (true)
        {
            Console.Write("Укажите путь к папке с Excel файлами: ");
            rootFolder = Console.ReadLine()?.Trim() ?? "";
            if (Directory.Exists(rootFolder)) break;
            Console.WriteLine("❌ Папка не найдена! Попробуйте ещё раз.\n");
        }

        // 2. Лист
        Console.Write("Введите название листа (оставьте пустым для сканирования всех листов): ");
        string sheetNameFilter = Console.ReadLine()?.Trim() ?? "";

        // 3. Ячейки
        Console.Write("Введите ячейки через запятую (например: A1,B1,C1,D1): ");
        var targetCells = (Console.ReadLine()?.Trim() ?? "")
                          .Split(',', StringSplitOptions.RemoveEmptyEntries)
                          .Select(c => c.Trim())
                          .ToList();

        if (targetCells.Count == 0)
            ExitWithMessage("Не указаны ячейки. Завершение программы.");

        // 4. Заголовки
        var results = new List<List<string>>();
        var header = new List<string> { "Файл (относительный путь)", "Имя файла", "Имя листа" };
        header.AddRange(targetCells);
        results.Add(header);

        // 5. Сканирование
        var excelFiles = Directory.GetFiles(rootFolder, "*.*", SearchOption.AllDirectories)
                                  .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                                           || f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)
                                           || f.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase))
                                  .ToList();

        int total = excelFiles.Count;
        int current = 0;

        foreach (var file in excelFiles)
        {
            current++;
            if (Path.GetFileName(file).StartsWith("~$"))
            {
                DrawProgressBar(current, total);
                continue;
            }

            try
            {
                if (file.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase))
                {
                    // ---------- ЧТЕНИЕ XLSB через ExcelDataReader ----------
                    using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var reader = ExcelReaderFactory.CreateReader(stream)) // авто-детект формата (в т.ч. .xlsb)
                    {
                        do
                        {
                            string sheetName = reader.Name ?? ""; // реальное имя листа
                            if (!string.IsNullOrEmpty(sheetNameFilter) &&
                                !sheetName.Equals(sheetNameFilter, StringComparison.OrdinalIgnoreCase))
                            {
                                // лист не подходит под фильтр — переходим к следующему
                                continue;
                            }

                            // читаем весь лист в память (построчно)
                            var sheetData = new List<object[]>();
                            while (reader.Read())
                            {
                                var values = new object[reader.FieldCount];
                                reader.GetValues(values);
                                sheetData.Add(values);
                            }

                            string relativePath = Path.GetRelativePath(rootFolder, file);
                            string fileNameOnly = Path.GetFileNameWithoutExtension(file);
                            var row = new List<string> { relativePath, fileNameOnly, sheetName };

                            foreach (var addr in targetCells)
                            {
                                string value = "";
                                try
                                {
                                    int col = ColLetterToIndex(GetColumnLetters(addr)); // 0-based
                                    int r = GetRowNumber(addr) - 1;                     // 0-based
                                    if (r >= 0 && r < sheetData.Count)
                                    {
                                        var line = sheetData[r];
                                        if (col >= 0 && col < line.Length)
                                            value = line[col]?.ToString() ?? "";
                                    }
                                }
                                catch { /* пусто */ }
                                row.Add(value);
                            }

                            results.Add(row);

                        } while (reader.NextResult()); // следующий лист
                    }
                }
                else
                {
                    // ---------- ЧТЕНИЕ XLSX/XLSM через ClosedXML ----------
                    using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var workbook = new XLWorkbook(stream))
                    {
                        var sheetsToScan = string.IsNullOrEmpty(sheetNameFilter)
                            ? workbook.Worksheets.ToList()
                            : workbook.Worksheets
                                       .Where(s => s.Name.Equals(sheetNameFilter, StringComparison.OrdinalIgnoreCase))
                                       .ToList();

                        if (!sheetsToScan.Any())
                        {
                            Console.WriteLine($"\nПропуск {file}: лист '{sheetNameFilter}' не найден.");
                            DrawProgressBar(current, total);
                            continue;
                        }

                        foreach (var ws in sheetsToScan)
                        {
                            string relativePath = Path.GetRelativePath(rootFolder, file);
                            string fileNameOnly = Path.GetFileNameWithoutExtension(file);

                            var row = new List<string> { relativePath, fileNameOnly, ws.Name };

                            foreach (var addr in targetCells)
                            {
                                string value;
                                try { value = ws.Cell(addr).GetFormattedString(); }
                                catch { value = ""; }
                                row.Add(value);
                            }

                            results.Add(row);
                        }
                    }
                }
            }
            catch (HeaderException hex)
            {
                Console.WriteLine($"\nПропуск {file}: файл зашифрован или повреждён ({hex.Message}).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОшибка чтения {file}: {ex.Message}");
            }

            DrawProgressBar(current, total);
        }

        // 6. Сохранение результата
        string timestamp = DateTime.Now.ToString("yyyy.MM.dd HH-mm");
        string resultFileName = $"{timestamp} Result.xlsx";
        string resultFilePath = Path.Combine(rootFolder, resultFileName);

        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.Worksheets.Add("Собранные данные");

            for (int i = 0; i < results.Count; i++)
                for (int j = 0; j < results[i].Count; j++)
                    ws.Cell(i + 1, j + 1).Value = results[i][j];

            ws.Row(1).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();
            workbook.SaveAs(resultFilePath);
        }

        Console.WriteLine($"\n\n✅ Готово! Все данные собраны в: {resultFilePath}");
        Console.WriteLine("Нажмите любую клавишу для выхода...");
        Console.WriteLine("\n\n (с) Галиев Ленар \n https://github.com/LEN4R/ScanAllExcelFiles");
        Console.ReadKey();
    }

    // --- utils ---
    static void ExitWithMessage(string message)
    {
        Console.WriteLine(message);
        Console.WriteLine("Нажмите любую клавишу для выхода...");
        Console.ReadKey();
        Environment.Exit(0);
    }

    static void DrawProgressBar(int progress, int total, int barSize = 50)
    {
        double percent = total == 0 ? 1 : (double)progress / total;
        int filled = (int)(percent * barSize);

        Console.CursorLeft = 0;
        Console.Write("[");
        Console.Write(new string('#', Math.Clamp(filled, 0, barSize)));
        Console.Write(new string('-', Math.Clamp(barSize - filled, 0, barSize)));
        Console.Write($"] {percent:P0} ({progress}/{total})");
    }

    static string GetColumnLetters(string cellAddress)
        => new string(cellAddress.Where(char.IsLetter).ToArray());

    static int GetRowNumber(string cellAddress)
    {
        var num = new string(cellAddress.Where(char.IsDigit).ToArray());
        return int.TryParse(num, out int r) ? r : -1;
    }

    static int ColLetterToIndex(string colLetters) // A->0, B->1, ... AA->26
    {
        int sum = 0;
        foreach (char c in colLetters.ToUpperInvariant())
        {
            sum = checked(sum * 26 + (c - 'A' + 1));
        }
        return sum - 1;
    }
}
