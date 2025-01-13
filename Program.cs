using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System;
using System.Diagnostics;
using System.Text;

namespace Excel2Json
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            if (args == null || args.Length <= 0)
            {
                Console.WriteLine("请指定json文件");
                return;
            }

            string jsonPath = Path.GetFullPath(args[0]);
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"json文件不存在：{jsonPath}");
                return;
            }

            string jsonContent = File.ReadAllText(jsonPath);
            CfgModel model = JsonConvert.DeserializeObject<CfgModel>(jsonContent);

            if (model == null)
            {
                Console.WriteLine($"配置数据类转换失败：{jsonPath}");
                return;
            }

            model.excelDir = Path.GetFullPath(model.excelDir);
            model.outputPath = Path.GetFullPath(model.outputPath);

            Console.WriteLine($"excel路径:{model.excelDir}");
            if (!Directory.Exists(model.excelDir))
            {
                Console.WriteLine($"excel路径不存在");
                return;
            }

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            try
            {
                Generate(model);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return;
            }

            stopwatch.Stop();
            TimeSpan elapsedTime = stopwatch.Elapsed;

            Console.WriteLine("代码执行时间： " + elapsedTime);
            Console.ReadLine();
        }

        private static void Generate(CfgModel model) 
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string output = model.outputPath;
            if (!Path.Exists(output))
            {
                Directory.CreateDirectory(output);
            }

            string[] paths = Directory.GetFiles(model.excelDir, "*.xlsx", SearchOption.AllDirectories);

            for (int i = 0; i < paths.Length; i++)
            {
                string fileName = Path.GetFileNameWithoutExtension(paths[i]);
                if (string.IsNullOrEmpty(fileName) || fileName.StartsWith('~')) 
                {
                    continue;
                }

                using (FileStream stream = new FileStream(Path.GetFullPath(paths[i]), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (ExcelPackage package = new ExcelPackage(stream))
                    {

                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        List<(string, string)> head = new List<(string, string)>();
                        Queue<(string, string)> values = new Queue<(string, string)>();

                        for (int row = 1; row <= rowCount; row++)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                if (worksheet.Cells[1, col].Value == null || worksheet.Cells[1, col].Value.ToString() == "#")
                                {
                                    continue;
                                }

                                if (row == 3)
                                {
                                    if (worksheet.Cells[row, col].Value == null)
                                    {
                                        continue;
                                    }

                                    string propName = worksheet.Cells[1, col].Value.ToString();
                                    propName = char.ToUpper(propName[0]) + propName[1..];

                                    string valueType = worksheet.Cells[row, col].Value.ToString().ToLower();
                                    head.Add((propName, valueType));
                                }

                                if (row > 4)
                                {
                                    string valueType = worksheet.Cells[3, col].Value.ToString().ToLower();
                                    string value;
                                    if (worksheet.Cells[row, col].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, col].Value.ToString()))
                                    {
                                        if (string.Equals(valueType, "int"))
                                        {
                                            value = "0";
                                        }
                                        else if (string.Equals(valueType, "float"))
                                        {
                                            value = "0";
                                        }
                                        else if (string.Equals(valueType, "bool"))
                                        {
                                            value = "0";
                                        }
                                        else
                                        {                                            
                                            value = string.Empty;
                                        }
                                    }
                                    else
                                    {
                                        value = worksheet.Cells[row, col].Value.ToString();
                                        if (valueType == "bool")
                                        {
                                            value = value == "0" ? "false" : "true";
                                        }
                                    }

                                    values.Enqueue((valueType, value));
                                }
                            }
                        }

                        GenerateJson(fileName, output, head, values);
                    }
                }

            }

        }

        private static void GenerateJson(string fileName, string output, List<(string, string)> head, Queue<(string, string)> values) 
        {
                
            int count = values.Count / head.Count;
            Console.WriteLine($"GenerateJson: {fileName}, count: {count}");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("{");

            for (int i = 0; i < count; i++)
            {
                for (int j = 0; j < head.Count; j++) 
                {
                    var v = values.Dequeue();

                    if (j == 0)
                    {

                        if (string.IsNullOrEmpty(v.Item2))
                        {
                            continue;
                        }

                        sb.AppendLine($"\t\"{v.Item2}\": {{");    
                    }

                    string end = (j < head.Count - 1) ? "," : string.Empty + (j == head.Count - 1 ? "\n\t}" : string.Empty) + (i < count - 1 ? "," : string.Empty);

                    string vt = v.Item1;
                    if (string.Equals(vt, "int"))
                    {
                        sb.AppendLine($"\t\t\"{head[j].Item1}\": {v.Item2}{end}");
                    }
                    else if (string.Equals(vt, "float")) 
                    {
                        sb.AppendLine($"\t\t\"{head[j].Item1}\": {v.Item2}{end}");
                    }
                    else if (string.Equals(vt, "bool"))
                    {
                        sb.AppendLine($"\t\t\"{head[j].Item1}\": {v.Item2}{end}");
                    }
                    else 
                    {
                        sb.AppendLine($"\t\t\"{head[j].Item1}\": \"{v.Item2}\"{end}");
                    }
                }
            }
            sb.AppendLine("}");

            using (FileStream fs = new FileStream(Path.Combine(output, $"{fileName}.json"), FileMode.Create))
            {
                var data = Encoding.Default.GetBytes(sb.ToString());
                
                fs.Write(data, 0, data.Length);
                fs.Flush();
                fs.Close();
            }
        }
    }
}
