using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace WikiGamesParser
{
    class WriteExcel
    {
        static List<Game>               games;
        static Dictionary<string, int>  dictPlatforms = new Dictionary<string, int>();
        static Dictionary<string, int>  dictEngines = new Dictionary<string, int>();
        static Dictionary<string, int>  dictGenres = new Dictionary<string, int>();
        static int                      maxCols;
        static ExcelWorksheet           worksheet;
        static ExcelPackage             package;

        static private void init(string _path, string _year, out string _completePath)
        {
            string fileName = "Game list - " + _year;
            _completePath = _path + "\\" + fileName + DateTime.Now.ToString(@"MM-dd-yyyy HH-mm") + ".xls";
            var file    = new FileInfo(_completePath);         
            package     = new ExcelPackage(file);
            worksheet   = package.Workbook.Worksheets.Add(fileName);
        }

        static private void writeHeader(List<String> _engines, List<String> _platforms, List<String> _genres, string _year)
        {
            int i = 4;
            worksheet.Cells[1, 1, 1, 1].Value   = _year;
            worksheet.Cells[2, 1, 3, 1].Merge   = true;
            worksheet.Cells[2, 1, 3, 1].Value   = "Name";
            worksheet.Cells[2, 2, 2, 3].Merge   = true;
            worksheet.Cells[2, 2, 2, 3].Value   = "Modes";
            worksheet.Cells[3, 2].Value         = "Single-player";
            worksheet.Cells[3, 3].Value         = "Multiplayer";

            foreach (string platform in _platforms)
            {
                worksheet.Cells[3, i].Value = platform;
                dictPlatforms.Add(platform, i);
                i++;
            }
            worksheet.Cells[2, 4, 2, i - 1].Merge     = true;
            worksheet.Cells[2, 4, 2, i - 1].Value   = "Platforms";
            int startEnginesIndex = i;
            foreach (string engine in _engines)
            {
                worksheet.Cells[3, i].Value = engine;
                dictEngines.Add(engine, i);
                i++;
            }
            worksheet.Cells[2, startEnginesIndex, 2, i - 1].Merge = true;
            worksheet.Cells[2, startEnginesIndex, 2, i - 1].Value = "Engines";
            int startIndexGenres = i;
            foreach (string genre in _genres)
            {
                worksheet.Cells[3, i].Value = genre;
                dictGenres.Add(genre, i);
                i++;
            }
            worksheet.Cells[2, startIndexGenres, 2, i - 1].Merge = true;
            worksheet.Cells[2, startIndexGenres, 2, i - 1].Value = "Genres";
            worksheet.Cells[2, i, 2, i].Value = "Releases";
            maxCols = i;
            worksheet.Cells[2, maxCols + 1, 2, maxCols + 1].Value = "No Data";
            maxCols += 1;
        }

        public static void open(string _path)
        {
            FileInfo fileOpen = new FileInfo(_path);
            if (fileOpen.Exists)
            {
                System.Diagnostics.Process.Start(_path);
            }
            else
            {
                Console.WriteLine("File doesn't exist");
            }
        }

        public static void write(List<Game> _games, GetData _data, string _filePath, string _year)
        {
            games               = _games;
            string completePath = "";

            init(_filePath, _year, out completePath);

            try
            {
                Console.WriteLine("Writing data...");
                writeHeader(_data.getListsForExcel()[0], _data.getListsForExcel()[1], _data.getListsForExcel()[2], _year);
                writeRow();
                worksheet.Column(1).AutoFit();
                worksheet.Column(maxCols - 1).Width = 15;
                package.Save();
                Console.WriteLine("All data saved! - " + completePath);
                Console.WriteLine("Do you want open this file? (y/n):");
                string isOpenFile = Console.ReadLine();
                if (isOpenFile.ToLower().Contains("y"))
                {
                    open(completePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        static private void writeRow()
        {
            try
            {
                int i = 4;
                foreach (var game in games)
                {
                    worksheet.Cells[i, 1].Style.Fill.PatternType = ExcelFillStyle.None;
                    worksheet.Cells[i, 1].Style.Font.Color.SetColor(Color.Blue);
                    worksheet.Cells[i, 1].Hyperlink = new Uri(game.Link);
                    worksheet.Cells[i, 1].Value     = game.Name;                  

                    if (game.Mode == null)
                    {
                        game.Mode = "";
                    }
                    else if (game.Mode.Contains("ingle") && game.Mode.Contains("ulti"))
                    {
                        worksheet.Cells[i, 2].Value = "V";
                        worksheet.Cells[i, 3].Value = "V";
                    }
                    else if (game.Mode.Contains("ingle") || game.Mode.Contains("1"))
                    {
                        worksheet.Cells[i, 2].Value = "V";
                    }
                    else if (game.Mode.Contains("ulti"))
                    {
                        worksheet.Cells[i, 3].Value = "V";
                    }
                    else { }

                    if (game.Platforms != null)
                    {
                        foreach (string platform in game.Platforms)
                        {
                            if (dictPlatforms.ContainsKey(platform))
                            {
                                worksheet.Cells[i, dictPlatforms[platform]].Value = "V";
                            }
                        }
                    }

                    if (game.Engine != null)
                    {
                        if (dictEngines.ContainsKey(game.Engine))
                        {
                            worksheet.Cells[i, dictEngines[game.Engine]].Value = "V";                        
                        }
                    }

                    if (game.Genres != null)
                    {
                        foreach (string genres in game.Genres)
                        {
                            if (dictGenres.ContainsKey(genres))
                            {
                                worksheet.Cells[i, dictGenres[genres]].Value = "V";
                            }
                        }
                    }
                    else
                    {
                        worksheet.Cells[i,2,i, maxCols].Style.Fill.PatternType = ExcelFillStyle.MediumGray;
                        worksheet.Cells[i, 2, i, maxCols].Style.Fill.BackgroundColor.SetColor(Color.Red);
                        worksheet.Cells[i, maxCols].Value = "V";                      
                    }
                    worksheet.Cells[i, maxCols - 1].Style.WrapText  = true;
                    worksheet.Cells[i, maxCols - 1].Value           = game.getReleases(false);                
                    i++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
