using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace word
{
    class Program
    {
        static void Main(string[] args)
        {
            var inputName = string.Empty;
            var files = new string[0];
            var content = new StringBuilder();

            bool flag1 = false, flag2 = false;
            while (!flag1)
            {
                Console.WriteLine("输入要汇总的Word文档的文件夹名称");
                inputName = Console.ReadLine();
                files = Directory.GetFiles(inputName);
                foreach (var item in files)
                {
                    if (Path.GetExtension(item) == ".doc")
                    {
                        try
                        {
                            var doc = new Document();
                            doc.LoadFromFile(item);
                            var tempParagraphs = doc.Document.GetText().Replace("\r", "").Split('\n');
                            for (int i = 0; i < tempParagraphs.Length; i++)
                            {
                                var p = tempParagraphs[i];
                                if (Regex.IsMatch(p, "^\\s*$"))
                                    continue;
                                Console.WriteLine($"{i}\t{p}");
                            }
                            flag1 = true;
                            break;
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (Path.GetExtension(item) == ".docx")
                    {
                        try
                        {
                            var tempDocument = DocX.Load(item);
                            var tempParagraphs = tempDocument.Paragraphs;
                            for (int i = 0; i < tempParagraphs.Count; i++)
                            {
                                var p = tempParagraphs[i];
                                if (Regex.IsMatch(p.Text, "^\\s*$"))
                                    continue;
                                Console.WriteLine($"{i}\t{p.Text}");
                            }
                            flag1 = true;
                            break;
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
            int start = 0, end = 0;
            while (!flag2)
            {
                try
                {
                    Console.WriteLine("\n输入要汇总的行数，如 10-20");
                    var number = Console.ReadLine();
                    start = Convert.ToInt32(number.Split('-')[0]);
                    end = Convert.ToInt32(number.Split('-')[1]);
                    flag2 = true;
                }
                catch (Exception)
                {
                }
            }
            foreach (var item in files)
            {
                if (File.Exists(item))
                {
                    if (Path.GetExtension(item) == ".doc")
                    {
                        var doc = new Document();
                        doc.LoadFromFile(item);
                        var paragraphs = doc.Document.GetText().Replace("\r", "").Split('\n');
                        for (int i = start; i <= end; i++)
                        {
                            var p = paragraphs[i];
                            content.Append($"{p},");
                        }
                        content.AppendLine();
                    }
                    else if (Path.GetExtension(item) == ".docx")
                    {
                        var document = DocX.Load(item);
                        var paragraphs = document.Paragraphs;
                        for (int i = start; i <= end; i++)
                        {
                            var p = paragraphs[i];
                            content.Append($"{p.Text},");
                        }
                        content.AppendLine();
                    }
                }
            }
            Console.WriteLine("汇总成功，请输入保存的文件名");
            var outputName = Console.ReadLine();
            File.WriteAllText($"{inputName}/{outputName}.csv", content.ToString(), Encoding.UTF8);

            Console.WriteLine($"保存完成\t{inputName}/{outputName}.csv\t按任意键退出");
            Console.ReadKey();
        }
    }
}
