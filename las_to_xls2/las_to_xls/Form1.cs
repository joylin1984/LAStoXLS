using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace las_to_xls
{
    public partial class Form1 : Form
    {
        public class Curve : object
        {
            public string name { get; }
            public double min_value { get; set; }
            public double max_value { get; set; }
            public string measure { get; }

            public Curve(string _name, double _min_value, double _max_value, string _measure)
            {
                name = _name;
                min_value = _min_value;
                max_value = _max_value;
                measure = _measure;
            }

            public Curve(string _name, string _measure)
            {
                name = _name;
                min_value = 0;
                max_value = 0;
                measure = _measure;
            }

            public override string ToString()
            {
                return "" + min_value + " - " + max_value + ";" + measure;
            }
        }

        char separator = ',';
        string[] files;
        List<string> bad_dirs;
        List<string> bad_files;

        public Form1()
        {
            InitializeComponent();
            string sep = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;
            if (sep.Length == 1)
                separator = sep[0];
            bad_dirs = new List<string>();
            bad_files = new List<string>();
        }

        private void reset()
        {
            bad_dirs.Clear();
            bad_files.Clear();
            listBox1.Items.Clear();
            progressBar1.Value = 0;
        }

        private void button1_Click(object sender, EventArgs e)  //change directory
        {
            reset();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                label1.Text = folderBrowserDialog1.SelectedPath;
                files = getFiles(folderBrowserDialog1.SelectedPath, "*.las").ToArray(); //search for *.las files
                if (bad_dirs.Count != 0)
                {
                    string message = "";
                    foreach (string dir in bad_dirs)
                        message += dir + "\n";
                    MessageBox.Show("Из данных папок файлы не получилось считать из-за недостаточных прав доступа: \n" + message);
                }
                listBox1.Items.AddRange(files.Select(s => s.Remove(0, folderBrowserDialog1.SelectedPath.Length)).ToArray());
                label2.Text = files.Length + " файл(ов)";
                backgroundWorker1.RunWorkerAsync(files);
            }
        }

        private List<string> getFiles(string path, string pattern)
        {
            List<string> res = new List<string>(Directory.GetFiles(path, pattern));
            foreach (var dir in Directory.GetDirectories(path))
            {
                try
                {
                    res.AddRange(getFiles(dir, pattern));
                }
                catch (Exception )
                {
                    bad_dirs.Add(dir);
                }
            }
            return res;
        }
        
        private bool isEqual(double a, double b)
        {
            if (Math.Abs(a - b) < 0.0005)
                return true;
            return false;
        }

        private string getValueFromString(string str)
        {
            try
            {
                return str.Split(':').Last().Trim(' ');
            }
            catch (Exception ) { }
            return "-";
        }

        private double getFirstDoubleFromString(string str)
        {
            try
            {
                return Convert.ToDouble(str.Split(' ', ':').Where(s => Regex.IsMatch(s, @"^[+-]?\d{1,}\.?\d{1,}$")).First().Replace('.', separator));
            }
            catch (Exception ex) { }
            return -1;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)  //parsing and adding info in excelapp
        {
            Excel.Application excelapp = new Excel.Application();
            excelapp.Workbooks.Add();
            Excel._Worksheet worksheet = (Excel.Worksheet)excelapp.ActiveSheet;
            int row_num = 1;
            worksheet.Cells[row_num, "A"] = "Путь к файлу";
            worksheet.Cells[row_num, "B"] = "Номер скважины";
            worksheet.Cells[row_num, "C"] = "Дата";
            worksheet.Cells[row_num, "D"] = "Кровля скважины";
            worksheet.Cells[row_num, "E"] = "Подошва скважины";
            worksheet.Cells[row_num, "F"] = "Кровля кривой";
            worksheet.Cells[row_num, "G"] = "Подошва кривой";
            worksheet.Cells[row_num, "H"] = "Название кривой";
            worksheet.Cells[row_num, "I"] = "Единицы измерения";
            row_num++;
            string[] lasFiles = e.Argument as string[];
            int count = 0;
            foreach (string file in lasFiles)
            {
                StreamReader sr = new StreamReader(file, Encoding.GetEncoding(1251));
                string tmp = "";
                try
                {
                    while (!(tmp = sr.ReadLine()).Contains("STRT")) ;
                    double first_depth = getFirstDoubleFromString(tmp);
                    while (!(tmp = sr.ReadLine()).Contains("STOP")) ;
                    double last_depth = getFirstDoubleFromString(tmp);
                    while (!(tmp = sr.ReadLine()).Contains("NULL")) ;
                    double NULL = getFirstDoubleFromString(tmp);
                    while (!(tmp = sr.ReadLine()).Contains("WELL")) ;
                    string number = getValueFromString(tmp);
                    while (!(tmp = sr.ReadLine()).Contains("DATE")) ;
                    string date = getValueFromString(tmp);
                    while (!(tmp = sr.ReadLine()).Contains("DEPT")) ;
                    List<Curve> curves = new List<Curve>();
                    while (!(tmp = sr.ReadLine()).Contains("~ASCII"))
                    {
                        string[] arr = tmp.Split('.', ':').Where(s => !string.IsNullOrEmpty(s)).ToArray();
                        curves.Add(new Curve(arr[0].Trim(' '), arr[1].Trim(' ')));
                    }
                    while (!sr.EndOfStream)
                    {
                        string[] numbers = sr.ReadLine().Split(' ').Where(s => !string.IsNullOrEmpty(s)).Select(s => s.Replace('.', separator)).ToArray();
                        double depth = Convert.ToDouble(numbers[0]);
                        for (int i = 1; i < numbers.Length; ++i)
                        {
                            double number_i = 0;
                            Double.TryParse(numbers[i], out number_i);
                            if (curves[i - 1].min_value == 0 && !isEqual(number_i, NULL))
                                curves[i - 1].min_value = depth;
                            if (curves[i - 1].min_value != 0 && curves[i - 1].max_value < depth && !isEqual(number_i, NULL))
                                curves[i - 1].max_value = depth;
                        }
                    }
                    worksheet.Cells[row_num, "A"] = file;
                    worksheet.Cells[row_num, "B"] = number;
                    worksheet.Cells[row_num, "C"] = (date == "") ? "-" : date;
                    worksheet.Cells[row_num, "D"] = (isEqual(first_depth, -1)) ? "-" : first_depth.ToString();
                    worksheet.Cells[row_num, "E"] = (isEqual(last_depth, -1)) ? "-" : last_depth.ToString();
                    foreach (Curve c in curves)
                    {
                        worksheet.Cells[row_num, "F"] = c.min_value;
                        worksheet.Cells[row_num, "G"] = c.max_value;
                        worksheet.Cells[row_num, "H"] = c.name;
                        worksheet.Cells[row_num, "I"] = c.measure;
                        row_num++;
                    }
                }
                catch (Exception)
                {
                    bad_files.Add(file);
                }
                sr.Close();
                count++;
                (sender as BackgroundWorker).ReportProgress((int)(count * 100 / files.Length));
                worksheet.Columns[2].AutoFit();
                worksheet.Columns[3].AutoFit();
                worksheet.Columns[4].AutoFit();
                worksheet.Columns[5].AutoFit();
                worksheet.Columns[6].AutoFit();
                worksheet.Columns[7].AutoFit();
                worksheet.Columns[8].AutoFit();
                worksheet.Columns[9].AutoFit();
            }
            excelapp.Visible = true;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (bad_files.Count != 0)
            {
                string message = "";
                foreach (string file in bad_files)
                    message += file + "\n";
                MessageBox.Show(bad_files.Count + " файл(ов) не удалось распознать:\n" + message, "Отчет об ошибках", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
