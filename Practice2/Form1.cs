using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Practice2
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// Коллекция ID студентов
        /// </summary>
        public List<int> IDs = new List<int>();
        /// <summary>
        /// Поле общего количества студентов
        /// </summary>
        public int stCap = 0;
        /// <summary>
        /// Поле общего количества отличников
        /// </summary>
        public int otlCap = 0;
        /// <summary>
        /// Поле временной переменной всех оценок студента
        /// </summary>
        public int tmpvse = 0;
        /// <summary>
        /// Поле временной переменной оценок "5" студента
        /// </summary>
        public int tmpotl = 0;
        /// <summary>
        /// Поле процента отличников
        /// </summary>
        public double perc = 0;
        /// <summary>
        /// Коллекция оценок студентов
        /// </summary>
        public List<string> Marks = new List<string>();
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Кнопка открытия окна "Об авторе"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button2_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }
        /// <summary>
        /// Кнопка открытия CSV файла
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, EventArgs e)
        {
            string rfname = @"C:\r.csv";
            OpenFileDialog open = new OpenFileDialog();
            open.InitialDirectory = "С:\\";
            open.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            open.FilterIndex = 1;
            open.Title = "Открыть файл";
            if (open.ShowDialog() == DialogResult.OK)
            {
                rfname = open.FileName;
                using (TextFieldParser fs = new TextFieldParser(rfname))
                {
                    fs.TextFieldType = FieldType.Delimited;
                    fs.SetDelimiters(",");
                    fs.ReadFields();
                    while (!fs.EndOfData)
                    {
                        string[] fields = fs.ReadFields();
                        IDs.Add(Convert.ToInt32(fields[0]));
                        Marks.Add(fields[5]);
                    }
                }
            }
            //Бизнес-логика проекта
            for(int i=1;i<IDs.Count;i++)
            {
                if (IDs[i] != IDs[i - 1])
                    stCap++;
            }
            for (int i = 1; i < IDs.Count; i++)
            {
                if (IDs[i] == IDs[i - 1])
                {
                    tmpvse++;
                    if ((Marks[i] == "90") || (Marks[i] == "91") || (Marks[i] == "92") || (Marks[i] == "93") || (Marks[i] == "94") || (Marks[i] == "95") || (Marks[i] == "96") || (Marks[i] == "97") || (Marks[i] == "98") || (Marks[i] == "99") || (Marks[i] == "100"))
                        tmpotl++;
                }
                if (IDs[i] != IDs[i - 1])
                {
                    if (tmpvse == tmpotl)
                        otlCap++;
                    tmpvse = 0;
                    tmpotl = 0;
                }
            }
            checkBox1.Checked = true;
            label1.Text = "Всего студентов: " + stCap.ToString();
            label2.Text = "Всего отличников: " + otlCap.ToString();
            perc = (Convert.ToDouble(otlCap) * 100) / Convert.ToDouble(stCap);
            label3.Text = "Процент отличников: " + perc.ToString();
        }
        /// <summary>
        /// Кнопка создания XLS и сохранения результатов
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button3_Click(object sender, EventArgs e)
        {
                    Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel не установлен!");
                        return;
                    }
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[1, 1] = "Кол-во студентов";
                    xlWorkSheet.Cells[1, 2] = "Кол-во отличников";
                    xlWorkSheet.Cells[1, 3] = "Процент отличников";
                    xlWorkSheet.Cells[2, 1] = stCap.ToString();
                    xlWorkSheet.Cells[2, 2] = otlCap.ToString();
                    xlWorkSheet.Cells[2, 3] = perc.ToString();
                    xlWorkBook.SaveAs("marks.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    MessageBox.Show("Файл marks.xls в папке проекта успешно создан!");
        }
    }
}
