using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace WikiGamesParser
{
    class WriteExcel
    {
        static Application              oXL;
        static _Workbook                oWB;
        static _Worksheet               oSheet;
        static List<Game>               games;
        static Dictionary<string, int>  dictPlatforms = new Dictionary<string, int>();
        static Dictionary<string, int>  dictEngines  = new Dictionary<string, int>();
        static Dictionary<string, int>  dictGenres = new Dictionary<string, int>();
        static int                      maxCols;

        public static void write(List<Game> _games, GetData _data, string _filePath, string year)
        {
            games = _games;
            init();
            try
            {
                writeHeader(_data.getListsForExcel()[0], _data.getListsForExcel()[1], _data.getListsForExcel()[2], year);
                writeRow();
                oXL.UserControl = true;
                oWB.SaveAs(_filePath + "\\WikiParser_result1.xls", XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Console.WriteLine("All data saved");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static private void writeHeaderSimple()
        {
            oSheet.Cells[1, 1] = "Name";
            oSheet.Cells[1, 2] = "Releases";
            oSheet.Cells[1, 3] = "Platforms";
            oSheet.Cells[1, 4] = "Engine";
            oSheet.Cells[1, 5] = "Mode";
            oSheet.Cells[1, 6] = "Genre";
            oSheet.get_Range("A1", "F1").Font.Bold = true;
        }

        static private void writeHeader(List<String> engines, List<String> platforms, List<String> genres, string year)
        {
            int i = 4;
            oSheet.get_Range("A1", "A1").Value = year;
            oSheet.get_Range("A2", "A3").Merge();
            oSheet.get_Range("A2", "A3").Interior.Color = XlRgbColor.rgbAzure;
            oSheet.get_Range("A2", "A3").Value = "Name";
            oSheet.get_Range("A2", "A3").Interior.Color = XlRgbColor.rgbSkyBlue;
            oSheet.get_Range("B2", "C2").Merge();
            oSheet.get_Range("B2", "C2").Value = "Modes";
            oSheet.get_Range("B3", "B3").Value = "Single-player";
            oSheet.get_Range("C3", "C3").Value = "Multiplayer";

            foreach (string platform in platforms)
            {
                oSheet.Cells[3, i] = platform;
                dictPlatforms.Add(platform, i);
                i++;
            }
            oSheet.Range[oSheet.Cells[2, 4], oSheet.Cells[2, i-1]].Merge();
            oSheet.Range[oSheet.Cells[2, 4], oSheet.Cells[2, i-1]] = "Platforms";
            int startEnginesIndex = i;
            foreach (string engine in engines)
            {
                oSheet.Cells[3, i] = engine;
                dictEngines.Add(engine, i);
                i++;
            }
            oSheet.Range[oSheet.Cells[2, startEnginesIndex], oSheet.Cells[2, i - 1]].Merge();
            oSheet.Range[oSheet.Cells[2, startEnginesIndex], oSheet.Cells[2, i - 1]] = "Engines";
            int startIndexGenres = i;
            foreach (string genre in genres)
            {
                oSheet.Cells[3, i] = genre;
                dictGenres.Add(genre, i);
                i++;
            }
            oSheet.Range[oSheet.Cells[2, startIndexGenres], oSheet.Cells[2, i - 1]].Merge();
            oSheet.Range[oSheet.Cells[2, startIndexGenres], oSheet.Cells[2, i - 1]] = "Genres";

            oSheet.Range[oSheet.Cells[2, i], oSheet.Cells[2, i]] = "Releases";
            maxCols = i;
            oSheet.Range[oSheet.Cells[2, maxCols+1], oSheet.Cells[2, maxCols+1]] = "No Data";
            maxCols += 1;
        }

        static private void init()
        {
            oXL = new Application();
            oXL.Visible = true;
            oWB = (_Workbook)(oXL.Workbooks.Add(""));
            oSheet = (_Worksheet)oWB.ActiveSheet;          
        }

        static private void writeModes()
        {
            int i = 4;
            foreach (var game in games)
            {              
                oSheet.Hyperlinks.Add(oSheet.get_Range("A" + i, "A" + i), game.Link, Type.Missing, "Click to go", game.Name);
                oSheet.Columns.AutoFit();
                oSheet.Rows.AutoFit();
            }
        }

        static private void writeRow()
        {
            try
            {
                int i = 4;
                foreach (var game in games)
                {
                   
                     oSheet.Hyperlinks.Add(oSheet.get_Range("A" + i, "A" + i), game.Link, Type.Missing, "Click to go", game.Name);
                     oSheet.Columns.AutoFit();

                    if (game.Mode == null)
                    {
                        game.Mode = "";
                    }
                    else if (game.Mode.Contains("ingle") && game.Mode.Contains("ulti"))
                    {
                        oSheet.get_Range("B" + i, "B" + i).Value = "V";                   
                        oSheet.get_Range("C" + i, "C" + i).Value = "V";                     
                    }
                    else if (game.Mode.Contains("ingle") || game.Mode.Contains("1"))
                    {
                        oSheet.get_Range("B" + i, "B" + i).Value = "V";
                    }
                    else if (game.Mode.Contains("ulti"))
                    {
                        oSheet.get_Range("C" + i, "C" + i).Value = "V";
                    }
                    else
                    { }

                    if (game.Platforms != null)
                    {
                        foreach (string platform in game.Platforms)
                        {
                            if (dictPlatforms.ContainsKey(platform))
                            {
                                oSheet.Cells[i, dictPlatforms[platform]] = "V";
                            }
                        }
                    }

                    if (game.Engine != null)
                    {                        
                            if (dictEngines.ContainsKey(game.Engine))
                            {
                                oSheet.Cells[i, dictEngines[game.Engine]] = "V";
                            }                       
                    }

                if (game.Genres != null)
                {
                    foreach (string genres in game.Genres)
                    {
                        if (dictGenres.ContainsKey(genres))
                        {
                            oSheet.Cells[i, dictGenres[genres]] = "V";
                        }
                    }
                }
                else
                {
                    oSheet.Cells[i, maxCols] = "V";
                    oSheet.Range[oSheet.Cells[i, 1], oSheet.Cells[i, maxCols]].Interior.Color = XlRgbColor.rgbRed;
                }
                oSheet.Cells[i,maxCols-1] = game.getReleases(false);
                i++; 
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static private void writeRowDel()
        {
            int i = 3;
            foreach(var game in games)
            {
                oSheet.Columns[2].ColumnWidth = 10;
                oSheet.Hyperlinks.Add(oSheet.get_Range("A"+i, "A"+i), game.Link, Type.Missing, "Click to go", game.Name);
                oSheet.Columns.AutoFit();
                oSheet.Rows.AutoFit();
                oSheet.get_Range("B"+i, "B"+i).Value = game.getReleases(false);
                oSheet.Columns[2].ColumnWidth = 10;
                oSheet.Rows.AutoFit();
                oSheet.get_Range("C" + i, "C" + i).Value = game.getPlatforms(false);
                oSheet.Columns.AutoFit();
                oSheet.Rows.AutoFit();
                oSheet.get_Range("D" + i, "D" + i).Value = game.Engine;
                oSheet.Columns.AutoFit();
                oSheet.Rows.AutoFit();
                oSheet.get_Range("E" + i, "E" + i).Value = game.Mode;
                oSheet.Columns.AutoFit();
                oSheet.Rows.AutoFit();
                oSheet.get_Range("F" + i, "F" + i).Value = game.getGeners(false);
                oSheet.Columns.AutoFit();
                oSheet.Rows.AutoFit();
                i++;
            }
           



        }

    }
}
