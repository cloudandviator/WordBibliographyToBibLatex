using System;
using System.IO;
using System.Xml;

namespace Bibliography
{
    class Program
    {
        static void Main(string[] args)
        {
            string SourcesFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + Path.DirectorySeparatorChar + @"Microsoft\Bibliography\Sources.xml";
            string ExecutionDirectory = AppDomain.CurrentDomain.BaseDirectory + Path.DirectorySeparatorChar;
            string BibFile;
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(SourcesFile);

            XmlNodeList sourcesNodes = xmlDocument.GetElementsByTagName("b:Sources");

            XmlNodeList listSourcesNodes =  ((XmlElement)sourcesNodes[0]).GetElementsByTagName("b:Source");
            Source x = new Source();

            foreach (XmlElement nodo in listSourcesNodes)
            {
                x.SourceType = nodo.GetElementsByTagName("b:SourceType").Item(0).InnerText;
                x.tag = nodo.GetElementsByTagName("b:Tag").Item(0).InnerText;
                x.title = nodo.GetElementsByTagName("b:Title").Item(0).InnerText;
                try
                {
                    x.year = nodo.GetElementsByTagName("b:Year").Item(0).InnerText;
                }
                catch
                {
                }
                x.author = nodo.GetElementsByTagName("b:Author").Item(0).InnerText;
                switch (x.SourceType)
                {
                    case "JournalArticle":
                        x.journaltitle = nodo.GetElementsByTagName("b:JournalName").Item(0).InnerText;
                        BibFile = ExecutionDirectory + "JournalArticle.bib";
                        using (System.IO.StreamWriter file = new System.IO.StreamWriter(BibFile, true))
                        {
                            file.WriteLine(x.ToString());
                        }
                        break;
                    case "BookSection":
                        x.publisher = nodo.GetElementsByTagName("b:Publisher").Item(0).InnerText;
                        BibFile = ExecutionDirectory + "BookSection.bib";
                        using (System.IO.StreamWriter file = new System.IO.StreamWriter(BibFile, true))
                        {
                            file.WriteLine(x.ToString());
                        }
                        break;
                    case "DocumentFromInternetSite":
                    case "InternetSite":
                        x.howpublished = nodo.GetElementsByTagName("b:URL").Item(0).InnerText;
                        try
                        {
                            string year = nodo.GetElementsByTagName("b:YearAccessed").Item(0).InnerText;
                            string month = nodo.GetElementsByTagName("b:MonthAccessed").Item(0).InnerText;
                            string day = nodo.GetElementsByTagName("b:DayAccessed").Item(0).InnerText;
                            x.note = string.Format("[Accesado {0}-{1}-{2}]", year, month, day);
                        }
                        catch
                        {

                        }
                        BibFile = ExecutionDirectory + "InternetSite.bib";
                        using (System.IO.StreamWriter file = new System.IO.StreamWriter(BibFile, true))
                        {
                            file.WriteLine(x.ToString());
                        }
                        break;
                }                
            }
        }
        /// <summary>
        /// Common data struture
        /// </summary>
        public class Source
        {
            public string SourceType;
            public string tag;
            public string author;
            public string title;
            public string howpublished;
            public string year;
            public string note;
            public string journaltitle;
            public string publisher;
            public override string ToString()
            {
                string rst = string.Empty;
                switch (SourceType)
                {
                    case "JournalArticle":
                        rst += "@article{";
                        rst += string.Format("{0},\r\n", tag);
                        rst += string.Format("author = {{{0}}},\r\n", author);
                        rst += string.Format("title = {{{0}}},\r\n", title);
                        rst += string.Format("journaltitle = {{{0}}},\r\n", journaltitle);
                        rst += string.Format("year = {{{0}}},\r\n", year);
                        break;
                    case "BookSection":
                        rst += "@book{";
                        rst += string.Format("{0},\r\n", tag);
                        rst += string.Format("author = {{{0}}},\r\n", author);
                        rst += string.Format("title = {{{0}}},\r\n", title);
                        rst += string.Format("publisher = {{{0}}},\r\n", publisher);
                        rst += string.Format("year = {{{0}}},\r\n", year);
                        break;
                    case "DocumentFromInternetSite":
                    case "InternetSite":
                        rst += "@misc{";
                        rst += string.Format("{0},\r\n", tag);
                        rst += string.Format("author = {{{0}}},\r\n", author);
                        rst += string.Format("title = {{{0}}},\r\n", title);
                        rst += string.Format("howpublished = {{{0}}},\r\n", howpublished);
                        rst += string.Format("year = {{{0}}},\r\n", year);
                        rst += string.Format("note = {{{0}}},\r\n", note);
                        break;
                    default:
                        rst += string.Format("{0},\r\n", tag);
                        rst += string.Format("author = {{{0}}},\r\n", author);
                        rst += string.Format("title = {{{0}}},\r\n", title);
                        rst += string.Format("howpublished = {{{0}}},\r\n", howpublished);
                        rst += string.Format("year = {{{0}}},\r\n", year);
                        rst += string.Format("note = {{{0}}},\r\n", note);
                        break;
                }
                rst += "}";
                return rst;
            }
        }
    }
}
