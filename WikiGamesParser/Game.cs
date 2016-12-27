using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WikiGamesParser
{
    class Game
    {
        public int    Id { get; set; }
        public string Name { get; set; }
        public string Link { get; set; }
        public List<string> Genres { get; set; }
        public string Engine { get; set; }
        public List<string> Platforms { get; set; }
        public List<DateTime> Release { get; set; }
        public string Mode { get; set; }

        public override string ToString()
        {
            return "Id: " + Id + "\n" +
                "Name: " + Name + "\n" +
                "Link: " + Link + "\n" +
                "Gener: " + getGeners() + "\n" +
                "Engine: " + Engine + "\n" +
                "Platform: " + getPlatforms() + "\n" +
                "Release(s): " + getReleases() + "\n" +
                "Mode: " + Mode + "\n";
        }

        public string getReleases(bool toConsole = true)
        {
            string result = "";
            string newLine = "\n";
            int i = 0;
            if (Release == null)
                return "";
            else
            {
                if (toConsole)
                {
                    newLine += "\t";
                }
                foreach (var date in Release)
                {
                    if (!toConsole)
                    {
                        if (i == 0)
                            newLine = "";
                        else
                            newLine = "\n";
                    }
                    result += newLine + date.ToString("dd/MM/yyyy");
                    i++;
                }
                return result;
            }
        }

     
        public string getGeners(bool toConsole = true)
        {
            string result = "";
            string newLine = "\n";
            int i = 0;
            if (Genres == null)
                return "";
            else
            {
                if (toConsole)
                {
                    newLine += "\t";
                }
                foreach (var genre in Genres)
                {
                    if (!toConsole)
                    {
                        if (i == 0)
                            newLine = "";
                        else
                            newLine = "\n";
                    }
                    result += newLine + genre;
                    i++;
                }
                return result;
            }
        }
        public string getPlatforms(bool toConsole = true)
        {
            string result = "";
            string newLine = "\n";
            int i = 0;
            if (Platforms == null)
                return "";
            else
            {
                if (toConsole)
                {
                    newLine += "\t";
                }
                foreach (var platform in Platforms)
                {
                    if (!toConsole)
                    {
                        if (i == 0)
                            newLine = "";
                        else
                            newLine = "\n";
                    }
                    result += newLine + platform;
                    i++;
                }
                return result;
            }
        }
    }
}
