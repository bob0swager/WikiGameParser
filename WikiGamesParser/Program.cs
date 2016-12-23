using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using HtmlAgilityPack;
using System.IO;
using System.Text.RegularExpressions;

namespace WikiGamesParser
{
    class Program
    {
        static string pageMainLink = @"https://en.wikipedia.org/wiki/";
        static int counter;
        static int withoutData;
        static List<Game> games = new List<Game>();
       
        static GetData data = new GetData();

        [STAThread]
        static void Main(string[] args)
        {
           string filePath = "c:\\";
            int year=0;
            while (year < 2004 || year > 2017)
            {
                Console.Write("Enter year of game: ");
                year = Convert.ToInt32(Console.ReadLine());
                if (year < 2004 || year > 2017)
                    Console.Write("The year is incorrect\n");
            }
            pageMainLink += year + "_in_video_gaming";
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            fbd.Description = "Select folder to save file";
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePath = fbd.SelectedPath;
            }
            getWikiTables();
            WriteExcel.write(games, data, filePath);
            Console.Read();
        }

        static void getWikiTables()
        {
            int countTables = 0;
            withoutData = 0;

            var Webget = new HtmlWeb();
            var doc = Webget.Load(pageMainLink);
            foreach (HtmlNode node in doc.DocumentNode.Descendants("table").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("wikitable")))
            {
                HtmlDocument newNode = new HtmlDocument();
                newNode.LoadHtml(node.InnerHtml);

                if (countTables == 5 || countTables == 6 || countTables == 7 || countTables == 8)
                {
                    foreach (HtmlNode gameLink in newNode.DocumentNode.SelectNodes("//tr//td//i//a").ToArray())
                    {
                        getCurrentPageCode(gameLink.Attributes["href"].Value, gameLink.InnerText);
                        counter++; 
                    }
                }
                countTables++;
               
            }
            Console.WriteLine("without data: %1", withoutData);
        }

        static void getCurrentPageCode(string gameLink, string gameName)
        {
            Game game = new Game();
            string http = "https://en.wikipedia.org";
            gameLink = http + gameLink;
            game.Id = counter;
            game.Link = gameLink;
            game.Name = gameName;

            using (WebClient client = new WebClient())
            {
                string html = null;
                try
                {
                    html = client.DownloadString(gameLink);
                    fillGameObject(html, game);
                }
                catch (Exception ex)
                {
                }
                Console.WriteLine("------------");
            }
            Console.WriteLine(game.ToString());
            games.Add(game);
        }

        static void fillGameObject(string html, Game game)
        {
            HtmlDocument newNode = new HtmlDocument();
            newNode.LoadHtml(html);
            try
            {
                HtmlNode tableData = newNode.DocumentNode.Descendants("table").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("infobox")).First();
                string parm_val = null;
                foreach (HtmlNode parameter in tableData.SelectNodes("//tr"))
                {
                    if (parameter.InnerText.Contains("Genre"))
                    {
                        parm_val = parameter.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(1);
                        game.Genres = data.getGenres(parm_val);
                    }
                    else if (parameter.InnerText.Contains("Developer"))
                    {
                        parm_val = parameter.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(1);
                        game.Developer = parm_val;
                    }
                    else if (parameter.InnerText.Contains("Designer"))
                    {
                        parm_val = parameter.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(1);
                        game.Designer = parm_val;
                    }
                    else if (parameter.InnerText.Contains("Engine"))
                    {
                        parm_val = parameter.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(1);
                        game.Engine = parm_val;
                        data.getEngines(parm_val);
                    }
                    else if (parameter.InnerText.Contains("Platform"))
                    {
                        int i = 0;
                        parm_val = "";
                        foreach (var platform in parameter.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)))
                            {
                            if(i!=0)
                                parm_val += platform + ", ";
                            i++;
                        }
                        game.Platforms = data.getPlatforms(parm_val);
                    }
                    else if (parameter.InnerText.Contains("Mode"))
                    {
                        if (game.Mode == null)
                        {
                            try
                            {
                                parm_val = parameter.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(1);
                                if (String.IsNullOrEmpty(parm_val))
                                    parm_val = "";
                                game.Mode = parm_val;
                            }
                            catch (ArgumentOutOfRangeException ex)
                            {
                                continue;
                            }
                        }
                    }
                    else if (parameter.InnerText.Contains("Release"))
                    {
                        try
                        {
                            parm_val = "";
                            int c = 0;
                            foreach (string elem in parameter.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)))
                            {
                                if (!elem.Contains("Release") && Regex.IsMatch(elem, "[0-9]{2}"))
                                {
                                    parm_val += parameter.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(c) + "|";
                                }
                                c++;
                            }
                        }
                        catch (ArgumentOutOfRangeException ex)
                        {
                            continue;
                        }
                        game.Release = GetData.getDateReleases(parm_val);
                    }
                    else if (parameter.InnerText.Contains("Publisher"))
                    {
                        parm_val = parameter.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(1);
                        game.Publisher = parm_val;
                    }
                    else if (parameter.InnerText.Contains("Artist"))
                    {
                        parm_val = parameter.InnerText.Split('\n').Where(a => !string.IsNullOrEmpty(a)).ElementAt(1);
                        game.Artist = parm_val;
                    }
                    else
                    {

                    }
                }
            }
            catch (InvalidOperationException ex)
            {
                withoutData++;
            }
            int res = 0;           
        }
        
        static void writeToCSV(string title, string link)
        {
            var forbiddenChars = @",;:".ToCharArray();
            var csv = new StringBuilder();
            var newLine = string.Format("{0};{1}", new string(title.Where(c => !forbiddenChars.Contains(c)).ToArray()), new string(link.Where(c => !forbiddenChars.Contains(c)).ToArray()));
            csv.AppendLine(newLine);
            File.AppendAllText(@"D:\testIntegration\Link.csv", csv.ToString());
        }
    }
}
