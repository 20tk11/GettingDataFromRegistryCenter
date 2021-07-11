using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;

namespace ConsoleApp1
{
    public class Data
    {
        private string dokdeskr { get; set; }
        private string dokdate { get; set; }
        private string getdate { get; set; }
        private string regdate { get; set; }
        private string psl { get; set; }
        public Data(string dokdeskr, string dokdate, string getdate, string regdate, string psl)
        {
            this.dokdeskr = dokdeskr;
            this.dokdate = dokdate;
            this.getdate = getdate;
            this.regdate = regdate;
            this.psl = psl;
        }
        public override string ToString()
        {
            return String.Format("|{0,-130}|{1,10}|{2,10}|{3,10}|{4,10}|", this.dokdeskr, this.dokdate, this.getdate, this.regdate, this.psl);
        }
        public string GetData(int i)
        {
            if (i == 1)
            {
                return this.dokdeskr;
            }
            else if (i == 2)
            {
                return this.dokdate;
            }
            else if (i == 3)
            {
                return this.getdate;
            }
            else if (i == 4)
            {
                return this.regdate;
            }
            else
            {
                return this.psl;
            }
        }
        // Remaining implementation of Person class.
    }
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }
        public void Close()
        {
            wb.Close();
        }
        public void Save()
        {
            wb.Save();
        }
        public void SaveAs(string path)
        {
            wb.SaveAs();
        }
        public void WriteToCell(int i, int j, string text)
        {
            ws.Cells[i + 1, j + 1].Value2 = text;
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            string[] alllines = File.ReadAllLines(@"D:\Data.txt");
            Array.Sort(alllines);
            var newlines = alllines.Distinct().ToArray();
            Console.WriteLine(alllines.Length);
            Console.WriteLine(newlines.Length);

            for(int i = 0; i < newlines.Length; i++)
            {
                if (i + 39 >= newlines.Length)
                {
                    break;
                }
                newlines[i] = newlines[i + 39];
            }
            foreach (var x in newlines)
            {
                Console.WriteLine(x);
            }
            HtmlAgilityPack.HtmlWeb website = new HtmlAgilityPack.HtmlWeb();
            List<Data> newlist = new List<Data>();
            int page_count = 1;
            int temp = 0;
            int temps = 1;
            int excellength = 0;
            int countofgeturl = 0;
            string heading = "";
            int pagecounts = 0;
            Excel excel = new Excel(@"D:\Book1.xlsx", 1);
            for (int i = 0; i < newlines.Length; i++)
            {
                while (page_count >= temps)
                {
                    if (countofgeturl == 80)
                    {
                        break;
                    }
                    string link = "https://www.registrucentras.lt/jar/p/dok.php?kod=" + newlines[i] + "&pav=%2A&p=" + temps.ToString();
                    countofgeturl++;
                    HtmlAgilityPack.HtmlDocument document = website.Load(link);
                    var list = document.DocumentNode.SelectNodes("//table[@cellspacing='1']//tr//td");
                    var meh = document.DocumentNode.SelectNodes("//table[@cellspacing='0']//tr//td//b");
                    int count = 0;
                    string pagecount = "";

                    int counts = 0;
                    string data1 = "";
                    string data2 = "";
                    string data3 = "";
                    string data4 = "";
                    string data5 = "";

                    foreach (var content in meh)
                    {
                        if (counts == 2)
                        {
                            pagecount = content.InnerText.ToString();
                            pagecounts = System.Convert.ToInt32(pagecount);
                        }
                        counts++;
                    }
                    if (pagecounts % 25 == 0)
                    { 
                        page_count = (pagecounts / 25);
                    }
                    else
                    {
                        page_count = (pagecounts / 25) + 1;
                    }
                    foreach (var content in list)
                    {

                        if (count < 1 && temps == 1)
                        {
                            heading += content.InnerText.ToString();
                        }
                        if (count == 6)
                        {
                            newlist.Add(new Data(data1, data2, data3, data4, data5));
                            temp++;
                            count = 1;
                        }
                        if (count == 1)
                        {
                            data1 = content.InnerText.ToString();
                        }
                        if (count == 2)
                        {
                            data2 = content.InnerText.ToString();
                        }
                        if (count == 3)
                        {
                            data3 = content.InnerText.ToString();
                        }
                        if (count == 4)
                        {
                            data4 = content.InnerText.ToString();
                        }
                        if (count == 5)
                        {
                            data5 = content.InnerText.ToString();
                        }
                        Console.WriteLine(content.InnerText);
                        Console.WriteLine(count);
                        count++;
                    }
                    newlist.Add(new Data(data1, data2, data3, data4, data5));
                    temps++;
                }
                excel.WriteToCell(excellength, 0, heading);
                excellength++;
                foreach (var x in newlist)
                {
                    excel.WriteToCell(excellength, 0, x.GetData(1));
                    excel.WriteToCell(excellength, 1, x.GetData(2));
                    excel.WriteToCell(excellength, 2, x.GetData(3));
                    excel.WriteToCell(excellength, 3, x.GetData(4));
                    excel.WriteToCell(excellength, 4, x.GetData(5));
                    excellength++;
                }
                int sss = 0;
                foreach(var x in newlist)
                {
                    sss++;
                }
                newlist.RemoveRange(0, sss);
                temps = 1;
                heading = "";
                //string filename = @"D:\temporary" + countofgeturl + ".xlsx";
                //excel.SaveAs(filename);
            }
            excel.Save();
            excel.Close();
            //Start();
        }
    }
}
