using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("=== Сканирование Excel файлов ===\n");

        // 1. Папка
        string rootFolder = "";
        while (true)
        {
            Console.Write("Укажите путь к папке с Excel файлами: ");
            rootFolder = Console.ReadLine()?.Trim() ?? "";

            if (Directory.Exists(rootFolder))
                break;

            Console.WriteLine("❌ Папка не найдена! Попробуйте ещё раз.\n");
        }

        // 2. Лист
        Console.Write("Введите название листа (оставьте пустым для сканирования всех листов): ");
        string sheetNameLine = Console.ReadLine()?.Trim() ?? "";

        // 3. Ячейки
        Console.Write("Введите ячейки через запятую (например: A1,B1,C1,D1): ");
        var targetCells = (Console.ReadLine()?.Trim() ?? "")
                          .Split(',', StringSplitOptions.RemoveEmptyEntries)
                          .Select(c => c.Trim())
                          .ToList();

        // Проверка
        if (targetCells.Count == 0)
        {
            ExitWithMessage("Не указаны ячейки. Завершение программы.");
        }

        // 4. Заголовки
        var results = new List<List<string>>();
        var header = new List<string> { "Файл (относительный путь)", "Имя файла", "Имя листа" };
        header.AddRange(targetCells);
        results.Add(header);

        // 5. Сканирование
        var excelFiles = Directory.GetFiles(rootFolder, "*.xlsx", SearchOption.AllDirectories);
        int total = excelFiles.Length;
        int current = 0;

        foreach (var file in excelFiles)
        {
            current++;
            try
            {
                using (var workbook = new XLWorkbook(file))
                {
                    List<IXLWorksheet> sheetsToScan;

                    if (string.IsNullOrEmpty(sheetNameLine))
                    {
                        // Пользователь не ввёл название → сканируем все листы
                        sheetsToScan = workbook.Worksheets.ToList();
                    }
                    else
                    {
                        var ws = workbook.Worksheets.FirstOrDefault(s => s.Name.Equals(sheetNameLine, StringComparison.OrdinalIgnoreCase));
                        sheetsToScan = ws != null ? new List<IXLWorksheet> { ws } : new List<IXLWorksheet>();
                    }

                    foreach (var ws in sheetsToScan)
                    {
                        string relativePath = Path.GetRelativePath(rootFolder, file);
                        string fileNameOnly = Path.GetFileNameWithoutExtension(file);

                        var row = new List<string> { relativePath, fileNameOnly, ws.Name };

                        foreach (var cell in targetCells)
                        {
                            string value = ws.Cell(cell).GetValue<string>();
                            row.Add(value);
                        }

                        results.Add(row);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОшибка чтения {file}: {ex.Message}");
            }

            DrawProgressBar(current, total);
        }

        // 6. Сохранение
        string timestamp = DateTime.Now.ToString("yyyy.MM.dd HH-mm");
        string resultFileName = $"{timestamp} Result.xlsx";
        string resultFilePath = Path.Combine(rootFolder, resultFileName);

        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.Worksheets.Add("Собранные данные");

            for (int i = 0; i < results.Count; i++)
            {
                for (int j = 0; j < results[i].Count; j++)
                {
                    ws.Cell(i + 1, j + 1).Value = results[i][j];
                }
            }

            ws.Row(1).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();

            workbook.SaveAs(resultFilePath);
        }

        Console.WriteLine($"\n\n Готово! Все данные собраны в: {resultFilePath}");
        Console.WriteLine("Нажмите любую клавишу для выхода...");
        Console.WriteLine("\n\n (с) Галиев Ленар \n https://github.com/LEN4R/ScanAllExcelFiles");
        Console.ReadKey();
    }

    static void ExitWithMessage(string message)
    {
        Console.WriteLine(message);
        Console.WriteLine("Нажмите любую клавишу для выхода...");
        Console.ReadKey();
        Environment.Exit(0);
    }

    static void DrawProgressBar(int progress, int total, int barSize = 50)
    {
        double percent = (double)progress / total;
        int filled = (int)(percent * barSize);

        Console.CursorLeft = 0;
        Console.Write("[");
        Console.Write(new string('#', filled));
        Console.Write(new string('-', barSize - filled));
        Console.Write($"] {percent:P0} ({progress}/{total})");
    }
}
