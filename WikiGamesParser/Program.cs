using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using HtmlAgilityPack;
using System.Text.RegularExpressions;

namespace WikiGamesParser
{
    class Program
    {      
        static List<Game>   games       = new List<Game>();
        static List<String> statistic   = new List<String>();
        static GetData      data        = new GetData();

        [STAThread]
        static void Main(string[] args)
        {
              int    i              = 0;
              string years          = ""; 
              string pathToFolder   = showFolderDialog();
              foreach (int year in checkYears())
              {               
                  getWikiTables(year);
                  if(i == 0)
                  {
                      years += year;
                  }
                  else
                      years += ", " + year;
                  i++;
              }
              showStatistic();
              WriteExcel.write(games, data, pathToFolder, years);
              Console.Read();          
        }

        static string showFolderDialog()
        {
            string outputFilePath = "c:\\";
            System.Windows.Forms.FolderBrowserDialog targetFolder = new System.Windows.Forms.FolderBrowserDialog();
            targetFolder.Description = "Select folder to save file";
            if (targetFolder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputFilePath = targetFolder.SelectedPath;
            }
            return outputFilePath;
        }

        static List<Int32> checkYears()
        {
            string  years       = "";
            int     firstYear   = 0,  
                    lastYear    = 0;
            List<Int32> yearsList = new List<Int32>() { };
            while (firstYear < 2004 || lastYear > 2017)
            {
                try
                {
                    Console.Write("Enter year/years of game (e.g 2014-2016 / 2012,2015): ");
                    years = Console.ReadLine();
                    yearsList = separateYears(years);
                    firstYear = yearsList[0];
                    lastYear = yearsList[yearsList.Count - 1];
                    if (firstYear < 2004 || lastYear > 2017)
                        Console.Write("The year/years is incorrect\n");
                }
                catch (Exception)
                {
                    Console.Write("The year/years is incorrect\n");
                }
            }             
            return yearsList;
        }

        static void showStatistic()
        {
            foreach (string line in statistic)
                Console.WriteLine(line);
        }

        static List<Int32> separateYears(string years)
        {
            List<Int32> years_result = new List<Int32>();
            if (years.Contains('-'))
            {
                foreach(string year in years.Split('-').ToArray())
                {
                    years_result.Add(Int32.Parse(new string(year.ToCharArray().Where(c => !Char.IsWhiteSpace(c)).ToArray())));
                }
                List<Int32> tmpYears_result = new List<Int32>();
                years_result.Sort();
                for (int i = years_result[0];i <= years_result[years_result.Count-1];i++)
                {        
                    tmpYears_result.Add(i);
                }
                years_result.Clear();
                years_result = tmpYears_result;
            }
            else if(years.Contains(','))
            {
                foreach (string year in years.Split(',').ToArray())
                {
                    years_result.Add(Int32.Parse(new string(year.ToCharArray().Where(c => !Char.IsWhiteSpace(c)).ToArray())));
                }
            }
            else if(Regex.IsMatch(years, @"^\d+$"))
            {
                years_result.Add(Int32.Parse(years));
            }
            years_result.Sort();
            return years_result;
        } 

        static void getWikiTables(int _year)
        {
            string  pageMainLink    = @"https://en.wikipedia.org/wiki/";
            int     countTables     = 0;
            int     gameCount       = 0;

            pageMainLink += _year + "_in_video_gaming";
            var Webget  = new HtmlWeb();
            var htmlDoc = Webget.Load(pageMainLink);
            foreach (HtmlNode node in htmlDoc.DocumentNode.Descendants("table").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("wikitable")))
            {
                HtmlDocument newNode = new HtmlDocument();
                newNode.LoadHtml(node.InnerHtml);          
                if (countTables == 5 || countTables == 6 || countTables == 7 || countTables == 8)
                {
                    if (newNode.DocumentNode.SelectSingleNode("//tr//td//i//a") != null)
                    {
                        foreach (HtmlNode gameLink in newNode.DocumentNode.SelectNodes("//tr//td//i//a").ToArray())
                        {
                            getCurrentPageCode(gameLink.Attributes["href"].Value, gameLink.InnerText, gameCount);
                            gameCount++;
                        }
                    }
                }
                countTables++;               
            }
            statistic.Add("All games in " + _year + " : " + gameCount);
        }

        static void getCurrentPageCode(string _link, string _name, int _number)
        {
            if (_name.Contains("Collection"))
            {
                string http1    = "https://en.wikipedia.org";
                _link           = http1 + _link;
                getCollectionGame(_link);
            }
            else
            {
                Game game   = new Game();
                string http = "https://en.wikipedia.org";
                _link       = http + _link;
                game.Id     = _number;
                game.Link   = _link;
                game.Name   = _name;

                using (WebClient client = new WebClient())
                {
                    string pamePageHtml = null;
                    try
                    {
                        pamePageHtml = client.DownloadString(_link);
                        fillGameObject(pamePageHtml, game);
                    }
                    catch (Exception) { }
                    Console.WriteLine("----------//----------");
                }
                Console.WriteLine(game.ToString());
                if(!games.Any(g => g.Name.Equals(_link)))
                {
                    games.Add(game);
                }
                    else
                {
                    compareGames(game);
                }
            }
        }

        static void compareGames(Game _game)
        {
            var existingGame = games.Find(item => item.Name == _game.Name);
            if (existingGame.Genres.SequenceEqual(_game.Genres))
            {
                foreach (string genre in existingGame.Genres)
                {
                    foreach (string new_genre in _game.Genres)
                    {
                        if (genre.Equals(new_genre))
                            if (!existingGame.Genres.Contains(new_genre))
                                existingGame.Genres.Add(new_genre);
                    }
                }
            }
            else if(!existingGame.Mode.Equals(_game.Mode))
            {
                existingGame.Mode += ", _game.Mode";
            }
            else if (existingGame.Platforms.SequenceEqual(_game.Platforms))
            {
                foreach (string platform in existingGame.Platforms)
                {
                    foreach (string new_platform in _game.Platforms)
                    {
                        if (platform.Equals(new_platform))
                            if (!existingGame.Platforms.Contains(new_platform))
                                existingGame.Platforms.Add(new_platform);
                    }
                }
            }
            else if (existingGame.Release.SequenceEqual(_game.Release))
            {
                foreach (DateTime dateTime in existingGame.Release)
                {
                    foreach (DateTime new_dateTime in _game.Release)
                    {
                        if (dateTime.Equals(new_dateTime))
                            if (!existingGame.Release.Contains(new_dateTime))
                                existingGame.Release.Add(new_dateTime);
                    }
                }
            }
        }

        static void getCollectionGame(string _link)
        {
            int gameCount = 0;
            List<Game> listGame = new List<Game>();
            var Webget = new HtmlWeb();
            var htmlDoc = Webget.Load(_link);
            foreach (HtmlNode node in htmlDoc.DocumentNode.Descendants("table").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("wikitable")))
            {
                HtmlDocument newNode = new HtmlDocument();
                newNode.LoadHtml(node.InnerHtml);
                try
                {
                    foreach (HtmlNode gameLink in newNode.DocumentNode.SelectNodes("//tr//td//i//a").ToArray())
                    {
                        getCurrentPageCode(gameLink.Attributes["href"].Value, gameLink.InnerText, gameCount);
                        gameCount++;
                    }
                }
                catch (ArgumentNullException)
                {
                    continue;
                }               
            }
        }

        static void fillGameObject(string _pageHtml, Game _game)
        {
            HtmlDocument mainNode = new HtmlDocument();
            mainNode.LoadHtml(_pageHtml);
            try
            {
                HtmlNode    tableData = mainNode.DocumentNode.Descendants("table").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("infobox")).First();
                string      parm_val    = null;

                foreach (HtmlNode param in tableData.SelectNodes("//tr"))
                {
                    if (param.InnerText.Contains("Genre"))
                    {
                        parm_val = param.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(1);
                        _game.Genres = data.getGenres(parm_val);
                    }
                    else if (param.InnerText.Contains("Engine"))
                    {
                        parm_val = param.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(1);
                        _game.Engine = parm_val;
                        data.getEngines(parm_val);
                    }
                    else if (param.InnerText.Contains("Platform"))
                    {
                        int i = 0;
                        parm_val = "";
                        foreach (var platform in param.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)))
                            {
                            if(i!=0)
                                parm_val += platform + ", ";
                            i++;
                        }
                        _game.Platforms = data.getPlatforms(parm_val);
                    }
                    else if (param.InnerText.Contains("Mode"))
                    {
                        if (_game.Mode == null)
                        {
                            try
                            {
                                parm_val = param.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(1);
                                if (String.IsNullOrEmpty(parm_val))
                                    parm_val = "";
                                _game.Mode = parm_val;
                            }
                            catch (ArgumentOutOfRangeException)
                            {
                                continue;
                            }
                        }
                    }
                    else if (param.InnerText.Contains("Release"))
                    {
                        try
                        {
                            parm_val = "";
                            int c = 0;
                            foreach (string elem in param.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)))
                            {
                                if (!elem.Contains("Release") && Regex.IsMatch(elem, "[0-9]{2}"))
                                {
                                    parm_val += param.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(c) + "|";
                                }
                                c++;
                            }
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            continue;
                        }
                        _game.Release = GetData.getDateReleases(parm_val);
                    }
                    else { }
                }
            }
            catch (InvalidOperationException){ }        
        }
    }
}
