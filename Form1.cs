using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;


namespace Gauss
{


    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
            round = Convert.ToInt32(numericUpDown1.Value);

        }
        string logs = "";
        bool cheeckErr = false;
        bool cheeck = false;
        bool Files = false; // Открытие файла / Opening Files
        int round; // Округление / Round
        int mode; // Номер dgv / Number dgv
        int size; // Размер матрицы / Size matrix
        int m;
        bool processing = false;
        
   
        public void ErrorFiles( string msg) // Функция создающая файл логов в случае его отсутствия или дописывает в конец файла / A function that creates a log file in case of its absence or appends to the end of the file
        {
            try
            {
                
                             
                    
                    File.AppendAllText("logs.txt", msg);
                    
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void dgvStyle (int index, int mode) // Форматирование таблицы / Formatting the table
        {
            switch (mode)
            {
                case 0:
                    {
                        dataGridView1.Columns[index].HeaderText = $"X{index + 1}";
                        dataGridView1.Columns[index].SortMode = DataGridViewColumnSortMode.NotSortable;
                        dataGridView1.Columns[index].Resizable = DataGridViewTriState.False;
                        dataGridView1.Rows[index].Resizable = DataGridViewTriState.False;
                        dataGridView1.Rows[index].HeaderCell.Value = $"{index + 1}";
                        break;
                    }
                case 1:
                    {
                        dataGridView2.Columns[index].HeaderText = $"X{index}";
                        dataGridView2.Columns[index].SortMode = DataGridViewColumnSortMode.NotSortable;
                        dataGridView2.Columns[index].Resizable = DataGridViewTriState.False;
                        dataGridView2.Rows[index - 1].Resizable = DataGridViewTriState.False;
                        dataGridView2.Rows[index - 1].HeaderCell.Value = $"{index}";
                        break;
                    }
            }
            

        } 

        public void CreatDataGridView( int size, int mode) //Создание таблицы / Creat the table
        {
            switch (mode)
            {
                case 0:
                    {
                        dataGridView1.ColumnCount = size + 1;
                        dataGridView1.RowCount = size;
                        dataGridView1.Columns[size].HeaderText = "B";
                        dataGridView1.Columns[size].SortMode = DataGridViewColumnSortMode.NotSortable;
                        dataGridView1.Columns[size].Resizable = DataGridViewTriState.False;

                        for (int i = 0; i < size; i++)
                        {
                            dgvStyle(i, mode);
                        }
                        break;
                    }
                case 1:
                    {
                        dataGridView2.ColumnCount = size + 2;
                        dataGridView2.RowCount = size;
                        dataGridView2.Columns[size+1].HeaderText = "B";
                        dataGridView2.Columns[size+1].SortMode = DataGridViewColumnSortMode.NotSortable;
                        dataGridView2.Columns[size+1].Resizable = DataGridViewTriState.False;
                        dataGridView2.Columns[0].HeaderText = "Ответ";
                        dataGridView2.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                        dataGridView2.Columns[0].Resizable = DataGridViewTriState.False;

                        for (int i = 1; i < size +1 ; i++)
                        {
                            dgvStyle(i, mode);
                        }
                        break;
                        
                    }
            }
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) // При смене значения combobox / When changing the combobox value
        {
            if (Files == false && processing == false)
            {
                mode = 0;
                size = comboBox1.SelectedIndex + 3; // Переменная хранит размер массива SIZExSIZE / The variable stores the size of the SIZExSIZE array
                CreatDataGridView(size, mode);

            }
           
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e) // Выход из приложения / Exiting the application
        {
            if (processing == false) {
                Application.Exit(); 
            }

        }


        

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e) // Сообщение о программе / Programm message
        {
            MessageBox.Show("Программа решает СЛАУ методом Гаусса\nПрограмма написана в ходе выполнения курсовой работы\nРазработчик: Диденко Максим Витальевич\nmvd2393123@gmail.com\nГУАП гр. 3045", "О программе", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                         
        }

        private void заполнитьМатрицуСлучайнымиЗначениямиToolStripMenuItem_Click(object sender, EventArgs e) // Заполнение матрицы случайными значениями / Filling the matrix with random value
        {
            if (processing == false) {
                Random r = new Random();
                Random random = new Random();
                Random drandom = new Random();
                int ch;
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    for (int j = 0; j < dataGridView1.RowCount; j++)
                    {
                        ch = r.Next(2);
                        switch (ch)
                        {
                            case 0:
                                dataGridView1.Rows[j].Cells[i].Value = Convert.ToDouble(random.Next(-100, 100));
                                if (Convert.ToDouble(dataGridView1.Rows[j].Cells[i].Value) == 0)
                                {
                                    dataGridView1.Rows[j].Cells[i].Value = 3045;
                                }
                                break;
                            case 1:
                                dataGridView1.Rows[j].Cells[i].Value = Math.Round(Convert.ToDouble(-100 + drandom.NextDouble() * (100 + 100)), round);
                                if (Convert.ToDouble(dataGridView1.Rows[j].Cells[i].Value) == 0)
                                {
                                    dataGridView1.Rows[j].Cells[i].Value = 3045;
                                }
                                break;
                        }
                    }
                }
            }
            }
        private void copyAlltoClipboard1() // Выделение всех объектов из dgv1 / Selecting all objects from dgv1
        {
            dataGridView1.SelectAll();
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText; 
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void copyAlltoClipboard2() // Выделение всех объектов из dgv2 / Selecting all objects from dgv2
        {
            dataGridView2.SelectAll();
            dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            DataObject dataObj = dataGridView2.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void сохранитьФайлСИсходнымиДаннымиToolStripMenuItem_Click(object sender, EventArgs e) // Сохранение файла с исходными данными в Excel / Saving a file with source data in Excel
        {
            if (processing == false) {
                this.Cursor = Cursors.WaitCursor;
                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                Excel.Worksheet ExcelWorksheet;
                try
                {
                    progressBar1.Value += 150;
                    copyAlltoClipboard1();
                    object misValue = System.Reflection.Missing.Value;
                    ExcelWorkBook = ExcelApp.Workbooks.Add(misValue);
                    ExcelWorksheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                    progressBar1.Value += 150;
                    Excel.Range CR = (Excel.Range)ExcelWorksheet.Cells[1, 1];
                    CR.Select();
                    progressBar1.Value += 200;
                    ExcelWorksheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    progressBar1.Value += 100;
                    Clipboard.Clear();
                    progressBar1.Value = 0;
                }
                catch (Exception ex)
                {
                    progressBar1.Value = 0;
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    ExcelApp.Visible = true;
                    ExcelApp.Quit();
                    this.Cursor = Cursors.Default;
                } 
            }
        }

      

        private void GaussDGV( int size) // Расчет по методу Гаусса / Calculation by the Gauss method
        {
            
            m = size + 1;
            
            int errRow = 0;
            double dig = 0;
            double digf = 0;
            int errCol = 0;
            bool variant3 = false;
            processing = true;

            
                try
                {
                    if (cheeckErr == false)
                    {
                        dataGridView2.Visible = false;
                    
                    progressBar1.Value += 150;
                    // Копирование из dgv1 в dgv2 / Copying from dgv1 to dgv2
                    for (int i = 0; i < size; i++)
                        {
                            for (int j = 0; j < size + 1; j++)
                            {
                                errCol = j;
                                errRow = i;
                                //dig += Convert.ToDouble(dataGridView1.Rows[i].Cells[i].Value);
                                dataGridView2.Rows[i].Cells[j + 1].Value = dataGridView1.Rows[i].Cells[j].Value;
                            }

                        }
                    

                    progressBar1.Value += 50;
                    

                    
                    
                        
                            progressBar1.Value += 200;
                        
                        // Прямой ход / Straight running
                        double tmp, delta = 0.0000001;
                        for (int columnIteration = 0; columnIteration < size; columnIteration++)
                        {

                            for (int i = columnIteration; i < size; i++)
                            {

                                tmp = Convert.ToDouble(dataGridView2.Rows[i].Cells[columnIteration + 1].Value);
                                if (Math.Abs(tmp) > delta)
                                {
                                    for (var j = 0; j < m; j++)
                                    {

                                        dataGridView2.Rows[i].Cells[j + 1].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells[j + 1].Value) / tmp;
                                    }
                                }

                                if (i != columnIteration)
                                {
                                    for (int j = 0; j < m; j++)
                                    {
                                        dataGridView2.Rows[i].Cells[j + 1].Value = (Convert.ToDouble(dataGridView2.Rows[i].Cells[j + 1].Value) - Convert.ToDouble(dataGridView2.Rows[columnIteration].Cells[j + 1].Value));
                                    }
                                }
                            }

                        }

                  


                  
                    progressBar1.Value += 50;
                       
                        
                        // Обратный ход / Reverse course
                        for (int i = 0; i < size; i++)
                        {
                            dataGridView2.Rows[i].Cells[0].Value = dataGridView2.Rows[i].Cells[m].Value;
                        }
                        for (int i = size - 2; i >= 0; i--)
                        {
                            for (int j = i + 1; j < size; j++)
                            {
                               
                                
                            
                                dataGridView2.Rows[i].Cells[0].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - Convert.ToDouble(dataGridView2.Rows[j].Cells[0].Value) * Convert.ToDouble(dataGridView2.Rows[i].Cells[j + 1].Value);
                            }
                            if (Math.Abs(Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) < delta)
                                dataGridView2.Rows[i].Cells[0].Value = 0;
                        }

                    }
                   
                        progressBar1.Value += 50;
                    
                    for (int i = 0; i < size; i++) // Округление значений / Rounding values
                    {
                        for (int j = 0; j < size + 2; j++)
                        {
                        
                       
                        dataGridView2.Rows[i].Cells[j].Value = Math.Round(Convert.ToDouble(dataGridView2.Rows[i].Cells[j].Value), round);
                        }
                       
                }

                

                double ans = 0;
                    for (int i = 0; i < size; i++)
                    {
                        ans += Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value);
                    }
                    if (ans == 0)
                    {
                       
                            progressBar1.Value = 0;

                            processing = false;
                            MessageBox.Show("Система не имеет решений", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        
                            dataGridView2.Visible = false;
                        dataGridView2.RowCount = 0;
                        
                        return;
                    }
                    
                        progressBar1.Value += 50;
                        cheeck = true;
                        dataGridView2.Visible = true;
                        progressBar1.Value = 0;
                        processing = false;

                    
                    
                }
                catch (Exception ex)
                {
                    
                        processing = false;
                        MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                        logs += Environment.NewLine + "*** Error! ***" + Environment.NewLine + "--------------" + Environment.NewLine + "ОШИБКА: " + ex.Message + Environment.NewLine + Environment.NewLine + "Метод: " + ex.TargetSite + Environment.NewLine + Environment.NewLine + "Вывод стека: " + ex.StackTrace + Environment.NewLine + Environment.NewLine + "Время возникновения: " + DateTime.Now + Environment.NewLine + Environment.NewLine;
                    ErrorFiles(logs);
                    dataGridView1.CurrentCell = dataGridView1.Rows[errRow].Cells[errCol];
                    dataGridView2.Visible = false;
                    progressBar1.Value = 0;
                    
                }
            
            
        }

   

    


        private void button1_Click(object sender, EventArgs e) // Кнопка "Решить" / Button "Решить"
        {
            if (processing == false) {
                mode = 1;
                this.Cursor = Cursors.WaitCursor;
                CreatDataGridView(size, mode);
                progressBar1.Value += 50;
                GaussDGV(size);
                this.Cursor = Cursors.Default;

            }
        }

        private void открытьИсходныйФайлToolStripMenuItem_Click(object sender, EventArgs e) // Октрытие файла Excel с исходными данными / Opening a Excel file with source data
        {
            if (processing == false)
            {

                try
                {
                    OpenFileDialog ofd = new OpenFileDialog();
                    ofd.DefaultExt = "*.xls;*.xlsx";
                    ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                    ofd.Title = "Выберите документ для загрузки данных";
                    if (ofd.ShowDialog() != DialogResult.OK)
                    {
                        progressBar1.Value = 0;
                        MessageBox.Show("Вы не выбрали файл для открытия", "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    
                        dataGridView1.RowCount = 0; // Удаление строк  / Deleting rows
                        dataGridView1.ColumnCount = 0; // Удаление столбцов // Deleting columns
                        Files = true;
                        this.Cursor = Cursors.WaitCursor;
                        progressBar1.Value += 50;
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            ofd.FileName + ";Extended Properties='Excel 12.0 XML;HDR=YES;IMEX=1';";

                        OleDbConnection con = new OleDbConnection(constr);
                        con.Open();
                        progressBar1.Value += 50;
                        DataSet ds = new DataSet(); // Формирование DataSet и DataTable / Formation of DataSet and DataTable
                        DataTable schemaTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                        progressBar1.Value += 50;
                        string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
                        string select = String.Format("SELECT * FROM [{0}]", sheet1);
                        OleDbDataAdapter ad = new OleDbDataAdapter(select, con);
                        ad.Fill(ds);
                        progressBar1.Value += 300;
                        DataTable dt = ds.Tables[0];
                        con.Close();
                        con.Dispose();
                        dataGridView1.DataSource = dt;
                        progressBar1.Value += 50;
                        size = dataGridView1.RowCount;
                        for (int i = 0; i < size; i++)
                        {
                            dgvStyle(i, 0);
                        }
                        progressBar1.Value += 100;
                        dt.Dispose();
                        progressBar1.Value = 0;
                        this.Cursor = Cursors.Default;

                   
                }

                catch (Exception ex)
                {
                    
                        progressBar1.Value = 0;
                        MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);


                        logs += Environment.NewLine + "*** Error! ***" + Environment.NewLine + "--------------" + Environment.NewLine + "ОШИБКА: " + ex.Message + Environment.NewLine + Environment.NewLine + "Метод: " + ex.TargetSite + Environment.NewLine + Environment.NewLine + "Вывод стека: " + ex.StackTrace + Environment.NewLine + Environment.NewLine + "Время возникновения: " + DateTime.Now + Environment.NewLine + Environment.NewLine;
                        ErrorFiles(logs);
                    
                }
            }

        }

        private void очистетьМатрицуToolStripMenuItem_Click(object sender, EventArgs e) // Удаление всех значений из ячейки / Deleting all values from a cell
        {
            if (processing == false) {
                try
                {
                    if (Files == true)
                    {
                        dataGridView1.DataSource = null;
                        dataGridView2.RowCount = 0;
                        dataGridView2.Visible = false;
                        Files = false;
                        size = 0;
                    }
                    else
                    {
                        dataGridView1.RowCount = 0;
                        dataGridView2.RowCount = 0;
                        dataGridView2.Visible = false;
                        dataGridView1.RowCount = size;
                    }
                    cheeck = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    logs += Environment.NewLine + "*** Error! ***" + Environment.NewLine + "--------------" + Environment.NewLine + "ОШИБКА: " + ex.Message + Environment.NewLine + Environment.NewLine + "Метод: " + ex.TargetSite + Environment.NewLine + Environment.NewLine + "Вывод стека: " + ex.StackTrace + Environment.NewLine + Environment.NewLine + "Время возникновения: " + DateTime.Now + Environment.NewLine + Environment.NewLine;
                    ErrorFiles(logs);
                }
            }


        }

        private void сохранитьФайлРешенияToolStripMenuItem_Click(object sender, EventArgs e) // Сохранение исходных данных и решения в файл Excel / Saving source data and solutions to an Excel file
        {
            if (cheeck == true && processing == false)
            {
                this.Cursor = Cursors.WaitCursor;
                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                Excel.Worksheet ExcelWorksheet;
                try
                {

                    progressBar1.Value += 50;
                    copyAlltoClipboard1();
                    object misValue = System.Reflection.Missing.Value;

                    ExcelWorkBook = ExcelApp.Workbooks.Add(misValue);
                    ExcelWorksheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                    Excel.Range CR = (Excel.Range)ExcelWorksheet.Cells[1, 1];
                    CR.Select();
                    progressBar1.Value += 150;
                    ExcelWorksheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    ExcelWorksheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
                    Clipboard.Clear();
                    progressBar1.Value += 150;
                    copyAlltoClipboard2();
                    ExcelWorksheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                    Excel.Range CR2 = (Excel.Range)ExcelWorksheet.Cells[1, 1];
                  
                    CR2.Select();
                    progressBar1.Value += 150;
                    ExcelWorksheet.PasteSpecial(CR2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);


                    progressBar1.Value += 100;
                    Clipboard.Clear();
                    progressBar1.Value = 0;
                }
                catch (Exception ex)
                {
                    
                    logs += Environment.NewLine + "*** Error! ***" + Environment.NewLine + "--------------" + Environment.NewLine + "ОШИБКА: " + ex.Message + Environment.NewLine + Environment.NewLine + "Метод: " + ex.TargetSite + Environment.NewLine + Environment.NewLine + "Вывод стека: " + ex.StackTrace + Environment.NewLine + Environment.NewLine + "Время возникновения: " + DateTime.Now + Environment.NewLine + Environment.NewLine;
                    ErrorFiles(logs);
                    progressBar1.Value = 0;
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {

                    ExcelApp.Visible = true;
                    ExcelApp.Quit();
                    this.Cursor = Cursors.Default;
                }
            }
            else
            {
                MessageBox.Show("Решение не было произведено", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e) // Число знаков после запятой / Number of decimal places
        {
            if (processing == false) {
                round = Convert.ToInt32(numericUpDown1.Value); 
            }

        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e) // Проверка вводимых данных и блокировка любых действий при некорректных данных / Checking the input data and blocking any actions in case of incorrect data 
        {
            try
            {
                const string disallowed = @"[^0-9-,]";
                var newText = Regex.Replace(e.FormattedValue.ToString(), disallowed, string.Empty);
                dataGridView1.Rows[e.RowIndex].ErrorText = "";
                if (dataGridView1.Rows[e.RowIndex].IsNewRow) return;
                if (string.CompareOrdinal(e.FormattedValue.ToString(), newText) == 0) return;
                e.Cancel = true;
                dataGridView1.Rows[e.RowIndex].ErrorText = "Некорректный символ!";
            }catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
                logs += Environment.NewLine + "*** Error! ***" + Environment.NewLine + "--------------" + Environment.NewLine + "ОШИБКА: " + ex.Message + Environment.NewLine + Environment.NewLine + "Метод: " + ex.TargetSite + Environment.NewLine + Environment.NewLine + "Вывод стека: " + ex.StackTrace + Environment.NewLine + Environment.NewLine + "Время возникновения: " + DateTime.Now + Environment.NewLine + Environment.NewLine;
                ErrorFiles(logs);
            }
        }

        private void справочныйМатериалToolStripMenuItem_Click(object sender, EventArgs e) // Открытие теоретического файла / Opening a theoretical file
        {
            if (processing == false)
            {
                try
                {
                    System.Diagnostics.Process.Start(@"Spravka.pdf");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    logs += Environment.NewLine + "*** Error! ***" + Environment.NewLine + "--------------" + Environment.NewLine + "ОШИБКА: " + ex.Message + Environment.NewLine + Environment.NewLine + "Метод: " + ex.TargetSite + Environment.NewLine + Environment.NewLine + "Вывод стека: " + ex.StackTrace + Environment.NewLine + Environment.NewLine + "Время возникновения: " + DateTime.Now + Environment.NewLine + Environment.NewLine;
                    ErrorFiles(logs);
                }
            }
        }

        private void руководствоToolStripMenuItem_Click(object sender, EventArgs e) // Открытие руководства / Opening the manual
        {
            if (processing == false)
            {
                try
                {
                    System.Diagnostics.Process.Start(@"Readme.pdf");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    logs += Environment.NewLine + "*** Error! ***" + Environment.NewLine + "--------------" + Environment.NewLine + "ОШИБКА: " + ex.Message + Environment.NewLine + Environment.NewLine + "Метод: " + ex.TargetSite + Environment.NewLine + Environment.NewLine + "Вывод стека: " + ex.StackTrace + Environment.NewLine + Environment.NewLine + "Время возникновения: " + DateTime.Now + Environment.NewLine + Environment.NewLine;
                    ErrorFiles(logs);
                }
            }
        }
    }
}

