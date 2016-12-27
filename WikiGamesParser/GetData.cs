using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WikiGamesParser
{
    class GetData
    {
        static List<String>         returnList;
        static public List<String>  engines      = new List<String>();
        static public List<String>  platforms    = new List<String>();
        static public List<String>  genres       = new List<String>();

        public List<String>[] getListsForExcel()
        {
            List<String>[] lists = new List<string>[3];
            lists[0] = engines;
            lists[1] = platforms;
            lists[2] = genres;
            return lists;
        }

        public void getEngines(string _engine)
        {
            if (_engine != null && _engine != "" && !engines.Contains(_engine))
            {
                string tmp_engine = "";
                tmp_engine = _engine;
                if (tmp_engine.Contains('('))
                {
                    tmp_engine = tmp_engine.Substring(0, tmp_engine.IndexOf('('));
                }
                else if (tmp_engine.Contains('['))
                {
                    tmp_engine = tmp_engine.Substring(0, tmp_engine.IndexOf('['));
                }
                if (!engines.Contains(tmp_engine) && !String.IsNullOrEmpty(tmp_engine))
                    engines.Add(tmp_engine);
            }          
        }
      
        public List<String> getPlatforms(string _platforms)
        {
            returnList = new List<String>();
            string tmp_platform = "";
            foreach (string platform in _platforms.Split(','))
            {
                tmp_platform = platform;
                if (tmp_platform[0] == ' ')
                    tmp_platform = tmp_platform.Substring(1);
                if (tmp_platform.Contains('('))
                {
                    tmp_platform = tmp_platform.Substring(0, tmp_platform.IndexOf('('));
                }
                else if (tmp_platform.Contains('['))
                {
                    tmp_platform = tmp_platform.Substring(0, tmp_platform.IndexOf('['));
                }
                returnList.Add(tmp_platform);
                if (!platforms.Contains(tmp_platform) && !String.IsNullOrEmpty(tmp_platform))
                    platforms.Add(tmp_platform);
            }
            return returnList;
        }

        public List<String> getGenres(string _genres)
        {
            returnList = new List<String>();
            string tmp_genre = "";
            foreach (string genre in _genres.Split(','))
            {
                tmp_genre = genre;
                if (tmp_genre[0] == ' ')
                    tmp_genre = tmp_genre.Substring(1);
                if (tmp_genre.Contains('('))
                {
                    tmp_genre = tmp_genre.Substring(0, tmp_genre.IndexOf('('));
                }
                else if(tmp_genre.Contains('['))
                {
                    tmp_genre = tmp_genre.Substring(0, tmp_genre.IndexOf('['));
                }
                returnList.Add(tmp_genre);
                if (!genres.Contains(tmp_genre))
                    genres.Add(tmp_genre);
            }
            return returnList;
        }

        public static List<DateTime> getDateReleases(string _dates)
        {
            Dictionary<string, int> months = new Dictionary<string, int>()
            {
                { "jan", 1},
                { "feb", 2},
                { "mar", 3},
                { "apr", 4},
                { "may", 5},
                { "jun", 6},
                { "jul", 7},
                { "aug", 8},
                { "sep", 9},
                { "oct", 10},
                { "nov", 11},
                { "dec", 12},
            };

            List<DateTime> releases = new List<DateTime>();
            Regex rgx = new Regex("[^0-9]");
            Regex rgx_number = new Regex("[^0-9]");
            string[] template1 = new String[5];
            foreach (string date in _dates.Split('|'))
            {
                DateTime date_w = new DateTime();
                string temp_date = "";
                int month = 1;
                int day = 1;
                int year = 1;
                temp_date = date;
                if (date.Contains(':'))
                {
                    temp_date = date.Substring(date.IndexOf(':') + 1);
                    if (temp_date[0] == ' ')
                    {
                        temp_date = temp_date.Substring(temp_date.IndexOf(temp_date[0]) + 1);
                    }
                }

                if (Regex.IsMatch(temp_date, "[0-9]{4}-[0-9]{2}-[0-9]{2}"))
                {
                    string[] template = temp_date.Substring((temp_date.Length - 11)).Split('-');

                    date_w = new DateTime(Int32.Parse(template[0]), Int32.Parse(template[1]), Int32.Parse(rgx.Replace(template[2], "")));
                    int j = 0;
                }
                else if (Regex.IsMatch(temp_date, "[0-9]{1,} [a-zA-Z]{3,} [0-9]{4}"))
                {
                    template1 = temp_date.Split(' ');
                    if (months.ContainsKey(template1[1].Substring(0, 3).Replace(" ", string.Empty).ToLower()))
                    {
                        month = months[template1[1].Substring(0, 3).ToLower()];
                    }
                    date_w = new DateTime(Int32.Parse(template1[2].Substring(0, 4)), month, Int32.Parse(template1[0]));
                    int df = 0;
                }
                else if (Regex.IsMatch(temp_date, "[0-9]{1,} [a-zA-Z]{3,}, [0-9]{4}"))
                {
                    template1 = temp_date.Split(' ');
                    if (months.ContainsKey(template1[1].Substring(0, 3).Replace(" ", string.Empty).ToLower()))
                    {
                        month = months[template1[1].Substring(0, 3).ToLower()];
                    }
                    date_w = new DateTime(Int32.Parse(template1[2]), month, Int32.Parse(template1[0]));
                }
                else if (Regex.IsMatch(temp_date, "[a-zA-Z]{3,} [0-9]{1,}, [0-9]{4}"))
                {
                    template1 = temp_date.Split(' ');
                    if (months.ContainsKey(template1[0].Substring(0, 3).Replace(" ", string.Empty).ToLower()))
                    {
                        month = months[template1[0].Substring(0, 3).ToLower()];
                    }
                    if (template1.Length >= 2)
                    {
                        day = Int32.Parse(rgx.Replace(template1[1], ""));
                        year = Int32.Parse(template1[2].Substring(0, 4));
                    }
                    date_w = new DateTime(year, month, day);
                }
                else if (Regex.IsMatch(temp_date, "[a-zA-Z]{3,} [0-9]{4}"))
                {
                    template1 = temp_date.Split(' ');
                    if (months.ContainsKey(template1[0].Substring(0, 3).Replace(" ", string.Empty).ToLower()))
                    {
                        month = months[template1[0].Substring(0, 3).ToLower()];
                    }
                    if (template1.Length >= 2)
                        year = Int32.Parse(template1[1]);
                    date_w = new DateTime(year, month, day);
                }
                else if (Regex.IsMatch(temp_date, "[a-zA-Z]{1}[1-4]{1} [0-9]{4}"))
                {
                    template1 = temp_date.Split(' ');
                    switch (template1[0][1])
                    {
                        case '1': month = 1; break;
                        case '2': month = 4; break;
                        case '3': month = 7; break;
                        case '4': month = 10; break;
                    };
                    year = Int32.Parse(template1[1].Substring(0, 4));
                    date_w = new DateTime(year, month, day);
                }
                else if (Regex.IsMatch(temp_date, "[0-9]{4}"))
                {
                    date_w = new DateTime(Int32.Parse(temp_date), month, day);
                    int df = 0;
                }
                else if (temp_date != "" && !temp_date.Contains("36"))
                {
                    int g = 0;
                }

                if (date_w > new DateTime(1, 1, 1))
                    if (!releases.Contains(date_w))
                        releases.Add(date_w);
            }
            return releases;
        }
    }
}
