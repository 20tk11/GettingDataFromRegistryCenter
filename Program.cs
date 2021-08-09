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
    class Job
    {
        public void Createtxt()
        {
            string[] alllines = File.ReadAllLines(@"D:\Projektai\Getting\Data1.txt");
            List<string> mylist = new List<string>();
            for (int i = 0; i < alllines.Length; i++)
            {
                if (!mylist.Contains(alllines[i]))
                {
                    mylist.Add(alllines[i]);
                }
            }
            using (StreamWriter outputFile = new StreamWriter(@"D:\Projektai\Getting\WriteLines.txt"))
            {
                foreach (string line in mylist)
                    outputFile.WriteLine(line);
            }
        }
        public void Filter()
        {
            string[] alllines = File.ReadAllLines(@"D:\Projektai\Getting\Data1.txt");
            string[] allliness = File.ReadAllLines(@"D:\Projektai\Getting\WriteLines.txt");
            List<int> mylist = new List<int>();
            for (int i = 0; i < allliness.Length; i++)
            {
                int count = 0;
                while(allliness[i] != alllines[count] && alllines.Length > count + 1 )
                {
                    //Console.WriteLine(count);
                    count++;
                }
                int place = count + 1;
                Console.WriteLine(i);
                Console.WriteLine(place);
                
                mylist.Add(place);
            }
            using (StreamWriter outputFile = new StreamWriter(@"D:\Projektai\Getting\WriteLines2.txt"))
            {
                foreach (int line in mylist)
                    outputFile.WriteLine(line);
            }
        }
        public void AssignCodes()
        {
            string[] alllines = File.ReadAllLines(@"D:\Projektai\Getting\Codes.txt");
            string[] allliness = File.ReadAllLines(@"D:\Projektai\Getting\WriteLines2.txt");
            using (StreamWriter outputFile = new StreamWriter(@"D:\Projektai\Getting\rezdata.txt"))
            {
                for (int i = 0; i < allliness.Length; i++)
                {
                    int tt = Int32.Parse(allliness[i]);
                    if (tt < alllines.Length)
                    {
                        Console.WriteLine(tt - 1);
                        outputFile.WriteLine(alllines[tt - 1]);
                    }
                }
            }
        }
        public void Do()
        {
            string[] alllines = File.ReadAllLines(@"D:\Projektai\Getting\rezdata.txt");
            HtmlAgilityPack.HtmlWeb website = new HtmlAgilityPack.HtmlWeb();
            List<Data> newlist = new List<Data>();
            int page_count = 1;
            int temp = 0;
            int temps = 1;
            int excellength = 0;
            int countofgeturl = 0;
            string heading = "";
            int pagecounts = 0;
            HtmlAgilityPack.HtmlDocument document;
            Excel excel = new Excel(@"D:\Book1.xlsx", 1);
            for (int i = 0; i < alllines.Length; i++)
            {
                while (page_count >= temps)
                {
                    if (countofgeturl == 140)
                    {
                        break;
                    }
                    string link = "https://www.registrucentras.lt/jar/p/dok.php?kod=" + alllines[i] + "&pav=%2A&p=" + temps.ToString();
                    countofgeturl++;
                    website.PreRequest = delegate (System.Net.HttpWebRequest webRequest)
                    {
                        webRequest.Timeout = 100000;
                        return true;
                    };
                    document = website.Load(link);
                    Console.WriteLine(link);
                    var list = document.DocumentNode.SelectNodes("//table[@cellspacing='1']//tr//td");
                    var meh = document.DocumentNode.SelectNodes("//table[@cellspacing='0']//tr//td[@width='20%']//b");
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
                        Console.WriteLine(content.InnerText.ToString());
                        if (counts == 0)
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
                foreach (var x in newlist)
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

    class Program
    {
        
        static void Main(string[] args)
        {
            Job job = new Job();
            job.Do();
        }
    }
}
