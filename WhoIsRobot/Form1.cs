using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PaddleOCRSharp;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SQLite;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing.Imaging;

namespace WhoIsRobot
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hwnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern void BlockInput(bool block);

        private string CDirectory = "";//当前目录
        private string DataDir = "";//数据目录
        private string DataFile = "";//数据文件
        private string taskLogFile = "";
        private static float MyDpi = 96F;
        private static double MyFontScale = 1;
        private bool CPUbool = true;
        private static PaddleOCREngine OCREngine = null;
        private static int smartWidth = 380;
        private static int smartHeight = 220;

        private static string GetFirstZiMu(string str)
        {
            if (string.Compare(str, "吖", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return str.ToLower();
            }
            if (string.Compare(str, "八", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "a";
            }
            if (string.Compare(str, "擦", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "b";
            }
            if (string.Compare(str, "耷", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "c";
            }
            if (string.Compare(str, "屙", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "d";
            }
            if (string.Compare(str, "发", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "e";
            }
            if (string.Compare(str, "旮", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "f";
            }
            if (string.Compare(str, "哈", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "g";
            }
            if (string.Compare(str, "丌", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "h";
            }
            if (string.Compare(str, "咔", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "j";
            }
            if (string.Compare(str, "垃", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "k";
            }
            if (string.Compare(str, "妈", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "l";
            }
            if (string.Compare(str, "拿", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "m";
            }
            if (string.Compare(str, "噢", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "n";
            }
            if (string.Compare(str, "趴", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "o";
            }
            if (string.Compare(str, "七", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "p";
            }
            if (string.Compare(str, "蚺", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "q";
            }
            if (string.Compare(str, "仨", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "r";
            }
            if (string.Compare(str, "他", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "s";
            }
            if (string.Compare(str, "挖", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "t";
            }
            if (string.Compare(str, "夕", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "w";
            }
            if (string.Compare(str, "丫", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "x";
            }
            if (string.Compare(str, "匝", false, CultureInfo.GetCultureInfo(2052)) < 0)
            {
                return "y";
            }
            if (string.Compare(str, "酢", false, CultureInfo.GetCultureInfo(2052)) <= 0)
            {
                return "z";
            }
            return str.ToLower();
        }

        private static string GetAllZM(string s)
        {
            string str = "";
            for (int i = 0; i < s.Length; i++)
            {
                str += GetFirstZiMu(s.Substring(i, 1));
            }
            return str;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                OCREngine = new PaddleOCREngine(null, new OCRParameter());
            }
            catch (Exception ex)
            {
                CPUbool = false;
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            UnregisterHotKey(this.Handle, 1578);
            UnregisterHotKey(this.Handle, 1579);
            RegisterHotKey(this.Handle, 1578, 0, (int)Keys.F8);
            RegisterHotKey(this.Handle, 1579, 0, (int)Keys.F9);
            using (Graphics gh = Graphics.FromHwnd(IntPtr.Zero))
            {
                MyDpi = gh.DpiX;
                if (MyDpi > 96F)
                {
                    Font MyNewFont = new Font("Microsoft YaHei UI", 9F);
                    this.Font = MyNewFont;
                    MyFontScale = (this.FontHeight) / 16D;
                    tabControl1.ItemSize = new Size((int)(tabControl1.ItemSize.Width * MyFontScale), (int)(tabControl1.ItemSize.Height * MyFontScale));
                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = MyNewFont;
                    dataGridView1.DefaultCellStyle.Font = MyNewFont;
                    dataGridView1.RowTemplate.Height = (int)(dataGridView1.RowTemplate.Height * MyFontScale);
                    dataGridView2.ColumnHeadersDefaultCellStyle.Font = MyNewFont;
                    dataGridView2.DefaultCellStyle.Font = MyNewFont;
                    dataGridView2.RowTemplate.Height = (int)(dataGridView2.RowTemplate.Height * MyFontScale);
                    dataGridView2.Columns[0].MinimumWidth = (int)(dataGridView2.Columns[0].MinimumWidth * MyFontScale);
                    int i1 = dataGridView2.ColumnCount;
                    for (int i = 2; i < i1; i++)
                    {
                        dataGridView2.Columns[i].MinimumWidth = (int)(dataGridView2.Columns[i].MinimumWidth * MyFontScale);
                    }
                    dataGridView3.ColumnHeadersDefaultCellStyle.Font = MyNewFont;
                    dataGridView3.DefaultCellStyle.Font = MyNewFont;
                    dataGridView3.RowTemplate.Height = (int)(dataGridView3.RowTemplate.Height * MyFontScale);
                    dataGridView4.ColumnHeadersDefaultCellStyle.Font = MyNewFont;
                    dataGridView4.DefaultCellStyle.Font = MyNewFont;
                    dataGridView4.RowTemplate.Height = (int)(dataGridView4.RowTemplate.Height * MyFontScale);
                    i1 = dataGridView4.ColumnCount;
                    for (int i = 0; i < i1; i++)
                    {
                        dataGridView4.Columns[i].MinimumWidth = (int)(dataGridView4.Columns[i].MinimumWidth * MyFontScale);
                    }
                    label1.Location = new Point(label1.Location.X, (int)(button8.Location.Y + button8.Height / 2D - label1.Height / 2D));
                    label11.Location = new Point(label11.Location.X, (int)(textBox7.Location.Y + textBox7.Height / 2D - label11.Height / 2D));
                    label12.Location = new Point(label12.Location.X, (int)(textBox8.Location.Y + textBox8.Height / 2D - label12.Height / 2D));
                    int i9 = (int)(button9.Location.Y + button9.Height / 2D - label2.Height / 2D);
                    label2.Location = new Point(label2.Location.X, i9);
                    label9.Location = new Point(label9.Location.X, i9);
                    label10.Location = new Point(label10.Location.X, i9);
                }
                else
                {
                    Font MyNewFont = new Font("SimSun", 9F);
                    this.Font = MyNewFont;
                    contextMenuStrip1.Font = MyNewFont;
                }
            }
            CDirectory = Directory.GetCurrentDirectory();
            if (CDirectory.Length > 3)
            {
                CDirectory += "\\";
            }
            DataDir = CDirectory + "Data\\";
            if (!Directory.Exists(DataDir))
            {
                Directory.CreateDirectory(DataDir);
            }
            DataFile = DataDir + "data.dat";
            if (!File.Exists(DataFile))
            {
                SQLiteConnection.CreateFile(DataFile);
                using (SQLiteConnection MySCon = new SQLiteConnection("Data Source=" + DataFile))
                {
                    MySCon.Open();
                    using (SQLiteCommand MySCommand = new SQLiteCommand(MySCon))
                    {
                        MySCommand.CommandText = "CREATE TABLE `任务` (`任务名称` TEXT NOT NULL UNIQUE);";
                        MySCommand.ExecuteNonQuery();
                        MySCommand.CommandText = "CREATE TABLE `任务设置` ("
                            + "`任务名称` TEXT,"
                            + "`序号` INT,"
                            + "`关键词` TEXT,"
                            + "`位次` INT,"
                            + "`位置X` INT,"
                            + "`位置Y` INT,"
                            + "`K1` TEXT,"
                            + "`K2` TEXT,"
                            + "`K3` TEXT,"
                            + "`延时` INT,"
                            + "`截屏` TEXT);";
                        MySCommand.ExecuteNonQuery();
                    }
                }
            }
            else
            {
                using (SQLiteConnection MySCon = new SQLiteConnection("Data Source=" + DataFile))
                {
                    MySCon.Open();
                    using (SQLiteCommand MySCommand = new SQLiteCommand(MySCon))
                    {
                        using (SQLiteDataAdapter MySDA = new SQLiteDataAdapter(MySCommand))
                        {
                            MySCommand.CommandText = "select 任务名称 from 任务";
                            MySDA.Fill(任务DataTable);
                        }
                    }
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                执行内容DataTable.Rows.Clear();
                this.Enabled = false;
                int fileCount = openFileDialog1.FileNames.Length;
                ConcurrentDictionary<string, int> loadCD = new ConcurrentDictionary<string, int>();
                try
                {
                    if (openFileDialog1.FilterIndex == 1)
                    {
                        for (int x = 0; x < fileCount; x++)
                        {
                            using (FileStream FS = File.OpenRead(openFileDialog1.FileNames[x]))
                            {
                                using (BinaryReader BR = new BinaryReader(FS))
                                {
                                    IWorkbook IWB;
                                    if (string.Compare(Encoding.ASCII.GetString(BR.ReadBytes(2)), "PK", StringComparison.OrdinalIgnoreCase) == 0)
                                    {
                                        FS.Seek(0, SeekOrigin.Begin);
                                        IWB = new XSSFWorkbook(FS);
                                    }
                                    else
                                    {
                                        FS.Seek(0, SeekOrigin.Begin);
                                        IWB = new HSSFWorkbook(FS);
                                    }
                                    ISheet MySheet = IWB.GetSheetAt(0);
                                    if (MySheet != null)
                                    {
                                        int MaxIndex = MySheet.LastRowNum;
                                        if (MaxIndex > 0)
                                        {
                                            for (int i = 0; i <= MaxIndex; i++)
                                            {
                                                string str1 = "", str2 = "", str3 = "", str4 = "", str5 = "";
                                                IRow MyRow = MySheet.GetRow(i);
                                                if (MyRow != null)
                                                {
                                                    try
                                                    {
                                                        if (MyRow.GetCell(0) != null)
                                                        {
                                                            if (MyRow.GetCell(0).CellType == CellType.Numeric)
                                                            {
                                                                str1 = MyRow.GetCell(0).NumericCellValue.ToString();
                                                            }
                                                            else
                                                            {
                                                                str1 = MyRow.GetCell(0).StringCellValue.Trim();
                                                            }
                                                        }
                                                    }
                                                    catch
                                                    { }
                                                    try
                                                    {
                                                        if (MyRow.GetCell(1) != null)
                                                        {
                                                            if (MyRow.GetCell(1).CellType == CellType.Numeric)
                                                            {
                                                                str2 = MyRow.GetCell(1).NumericCellValue.ToString();
                                                            }
                                                            else
                                                            {
                                                                str2 = MyRow.GetCell(1).StringCellValue.Trim();
                                                            }
                                                        }
                                                    }
                                                    catch
                                                    { }
                                                    try
                                                    {
                                                        if (MyRow.GetCell(2) != null)
                                                        {
                                                            if (MyRow.GetCell(2).CellType == CellType.Numeric)
                                                            {
                                                                str3 = MyRow.GetCell(2).NumericCellValue.ToString();
                                                            }
                                                            else
                                                            {
                                                                str3 = MyRow.GetCell(2).StringCellValue.Trim();
                                                            }
                                                        }
                                                    }
                                                    catch
                                                    { }
                                                    try
                                                    {
                                                        if (MyRow.GetCell(3) != null)
                                                        {
                                                            if (MyRow.GetCell(3).CellType == CellType.Numeric)
                                                            {
                                                                str4 = MyRow.GetCell(3).NumericCellValue.ToString();
                                                            }
                                                            else
                                                            {
                                                                str4 = MyRow.GetCell(3).StringCellValue.Trim();
                                                            }
                                                        }
                                                    }
                                                    catch
                                                    { }
                                                    try
                                                    {
                                                        if (MyRow.GetCell(4) != null)
                                                        {
                                                            if (MyRow.GetCell(4).CellType == CellType.Numeric)
                                                            {
                                                                str5 = MyRow.GetCell(4).NumericCellValue.ToString();
                                                            }
                                                            else
                                                            {
                                                                str5 = MyRow.GetCell(4).StringCellValue.Trim();
                                                            }
                                                        }
                                                    }
                                                    catch
                                                    { }
                                                    if (str1 != "" || str2 != "" || str3 != "" || str4 != "" || str5 != "")
                                                    {
                                                        StringBuilder sb = new StringBuilder();
                                                        sb.Append(str1);
                                                        sb.Append(str2);
                                                        sb.Append(str3);
                                                        sb.Append(str4);
                                                        sb.Append(str5);
                                                        if (loadCD.TryAdd(sb.ToString(), 0))
                                                        {
                                                            object[] ob = new object[5];
                                                            ob[0] = str1;
                                                            ob[1] = str2;
                                                            ob[2] = str3;
                                                            ob[3] = str4;
                                                            ob[4] = str5;
                                                            执行内容DataTable.Rows.Add(ob);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (openFileDialog1.FilterIndex == 2)
                    {
                        for (int x = 0; x < fileCount; x++)
                        {
                            string s1 = openFileDialog1.FileNames[x];
                            if (new FileInfo(s1).Extension.ToLower() == ".txt")
                            {
                                using (StreamReader sr = new StreamReader(s1, true))
                                {
                                    string[] str1 = { "\r", "\n" };
                                    string[] str2 = { "\t" };
                                    string[] strTmp = sr.ReadToEnd().Split(str1, StringSplitOptions.RemoveEmptyEntries);
                                    int i1 = strTmp.Length;
                                    for (int y = 0; y < i1; y++)
                                    {
                                        string[] strRead = strTmp[y].Split(str2, StringSplitOptions.None);
                                        int i2 = strRead.Length;
                                        object[] ob = new object[5];
                                        for (int z = 0; z < i2; z++)
                                        {
                                            if (z < 5)
                                            {
                                                ob[z] = strRead[z];
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        StringBuilder sb = new StringBuilder();
                                        for (int m = 0; m < 5; m++)
                                        {
                                            if (ob[m] != null)
                                            {
                                                sb.Append(ob[m]);
                                            }
                                        }
                                        if (loadCD.TryAdd(sb.ToString(), 0))
                                        {
                                            string s2 = ob[2].ToString();
                                            if (s2 != "")
                                            {
                                                if (s2.Length == 8 && (s2.Contains("-") == false))
                                                {
                                                    ob[2] = s2.Substring(0, 4) + "-" + s2.Substring(4, 2) + "-" + s2.Substring(6, 2);
                                                }
                                            }
                                            执行内容DataTable.Rows.Add(ob);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                using (StreamReader sr = new StreamReader(s1, Encoding.Unicode))
                                {
                                    string[] str1 = { "\r", "\n" };
                                    string[] str2 = { "\t" };
                                    string[] strTmp = sr.ReadToEnd().Split(str1, StringSplitOptions.RemoveEmptyEntries);
                                    int i1 = strTmp.Length;
                                    for (int y = 0; y < i1; y++)
                                    {
                                        string[] strRead = strTmp[y].Split(str2, StringSplitOptions.None);
                                        int i2 = strRead.Length;
                                        object[] ob = new object[5];
                                        for (int z = 0; z < i2; z++)
                                        {
                                            if (z < 5)
                                            {
                                                ob[z] = strRead[z];
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        StringBuilder sb = new StringBuilder();
                                        for (int m = 0; m < 5; m++)
                                        {
                                            if (ob[m] != null)
                                            {
                                                sb.Append(ob[m]);
                                            }
                                        }
                                        if (loadCD.TryAdd(sb.ToString(), 0))
                                        {
                                            执行内容DataTable.Rows.Add(ob);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                label10.Text = 执行内容DataTable.Rows.Count.ToString();
                this.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s1 = textBox1.Text.Trim();
            if (s1 != "")
            {
                ArrayList al = new ArrayList();
                for (int i = 0; i < 任务DataTable.Rows.Count; i++)
                {
                    al.Add(任务DataTable.Rows[i][0].ToString());
                }
                if (!al.Contains(s1))
                {
                    using (SQLiteConnection MySCon = new SQLiteConnection("Data Source=" + DataFile))
                    {
                        MySCon.Open();
                        using (SQLiteCommand MySCommand = new SQLiteCommand(MySCon))
                        {
                            MySCommand.CommandText = "Insert Into 任务 (任务名称) Values ('" + s1.Replace("'", "''") + "')";
                            MySCommand.ExecuteNonQuery();
                        }
                    }
                    任务DataTable.Rows.Add(s1);
                    任务DataTable.AcceptChanges();
                }
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            任务设置DataTable.Rows.Clear();
            if (dataGridView1.CurrentRow != null)
            {
                int i1 = dataGridView1.CurrentRow.Index;
                if (i1 > -1)
                {
                    taskStr = 任务DataTable.Rows[i1][0].ToString();
                    textBox1.Text = taskStr;
                    using (SQLiteConnection MySCon = new SQLiteConnection("Data Source=" + DataFile))
                    {
                        MySCon.Open();
                        using (SQLiteCommand MySCommand = new SQLiteCommand(MySCon))
                        {
                            using (SQLiteDataAdapter MySDA = new SQLiteDataAdapter(MySCommand))
                            {
                                MySCommand.CommandText = "select 序号,关键词,位次,位置X,位置Y,K1,K2,K3,延时,截屏 from 任务设置 where 任务名称='" + taskStr.Replace("'", "''") + "' ORDER BY 序号 ASC";
                                MySDA.Fill(任务设置DataTable);
                            }
                        }
                    }
                    string logDir = CDirectory + "Logs";
                    if (!Directory.Exists(CDirectory + "Logs"))
                    {
                        Directory.CreateDirectory(logDir);
                    }
                    taskLogFile = logDir + "\\" + taskStr + "日志.txt";
                }
            }
        }

        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (任务DataTable.Rows.Count > 0)
            {
                dataGridView1.ContextMenuStrip = contextMenuStrip1;
            }
            else
            {
                dataGridView1.ContextMenuStrip = null;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int i1 = dataGridView1.CurrentRow.Index;
                if (i1 > -1)
                {
                    if (MessageBox.Show("确定要删除选定的任务吗？", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.OK)
                    {
                        File.Delete(taskLogFile);
                        string s1 = taskStr.Replace("'", "''");
                        using (SQLiteConnection MySCon = new SQLiteConnection("Data Source=" + DataFile))
                        {
                            MySCon.Open();
                            using (SQLiteCommand MySCommand = new SQLiteCommand(MySCon))
                            {
                                MySCommand.Transaction = MySCon.BeginTransaction();
                                MySCommand.CommandText = "delete from 任务设置 where 任务名称='" + s1 + "'";
                                MySCommand.ExecuteNonQuery();
                                MySCommand.CommandText = "delete from 任务 where 任务名称='" + s1 + "'";
                                MySCommand.ExecuteNonQuery();
                                MySCommand.Transaction.Commit();
                                任务DataTable.Rows.RemoveAt(i1);
                                任务DataTable.AcceptChanges();
                            }
                        }
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                if (dataGridView1.CurrentRow.Index > -1)
                {
                    object[] ob = new object[10];
                    ob[0] = 任务设置DataTable.Rows.Count + 1;
                    ob[2] = 1;
                    ob[5] = "(空)";
                    ob[6] = "(空)";
                    ob[7] = "(空)";
                    ob[8] = 300;
                    ob[9] = "一次";
                    任务设置DataTable.Rows.Add(ob);
                    任务设置DataTable.AcceptChanges();
                    save任务设置(1, taskStr);
                }
            }
        }

        private void save任务设置(int i_t, string s1)//1 是保存，2 是克隆
        {
            s1 = s1.Replace("'", "''");
            int i2 = 任务设置DataTable.Rows.Count;
            using (SQLiteConnection MySCon = new SQLiteConnection("Data Source=" + DataFile))
            {
                MySCon.Open();
                using (SQLiteCommand MySCommand = new SQLiteCommand(MySCon))
                {
                    MySCommand.Transaction = MySCon.BeginTransaction();
                    if (i_t == 1)
                    {
                        MySCommand.CommandText = "delete from 任务设置 where 任务名称='" + s1 + "'";
                        MySCommand.ExecuteNonQuery();
                    }
                    else if (i_t == 2)
                    {
                        MySCommand.CommandText = "Insert Into 任务 (任务名称) Values ('" + s1 + "')";
                        MySCommand.ExecuteNonQuery();
                    }
                    for (int i = 1; i <= i2; i++)
                    {
                        MySCommand.CommandText = "Insert Into 任务设置 (任务名称,序号) Values ('" + s1 + "'," + i + ")";
                        MySCommand.ExecuteNonQuery();
                    }
                    for (int i = 0; i < i2; i++)
                    {
                        int i_序号 = i + 1;
                        string str关键词 = 任务设置DataTable.Rows[i][1].ToString().Replace("'", "''");
                        if (str关键词 != "")
                        {
                            MySCommand.CommandText = "Update 任务设置 set 关键词= '" + str关键词 + "' where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                            MySCommand.ExecuteNonQuery();
                        }
                        MySCommand.CommandText = "Update 任务设置 set 位次= " + (int)任务设置DataTable.Rows[i][2] + " where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                        MySCommand.ExecuteNonQuery();
                        string str位置X = 任务设置DataTable.Rows[i][3].ToString();
                        string str位置Y = 任务设置DataTable.Rows[i][4].ToString();
                        if (str位置X != "" && str位置Y != "")
                        {
                            MySCommand.CommandText = "Update 任务设置 set 位置X= " + int.Parse(str位置X) + " where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                            MySCommand.ExecuteNonQuery();
                            MySCommand.CommandText = "Update 任务设置 set 位置Y= " + int.Parse(str位置Y) + " where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                            MySCommand.ExecuteNonQuery();
                        }
                        else
                        {
                            MySCommand.CommandText = "Update 任务设置 set 位置X = null where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                            MySCommand.ExecuteNonQuery();
                            MySCommand.CommandText = "Update 任务设置 set 位置Y = null where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                            MySCommand.ExecuteNonQuery();
                        }
                        MySCommand.CommandText = "Update 任务设置 set K1= '" + 任务设置DataTable.Rows[i][5].ToString() + "' where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                        MySCommand.ExecuteNonQuery();
                        MySCommand.CommandText = "Update 任务设置 set K2= '" + 任务设置DataTable.Rows[i][6].ToString() + "' where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                        MySCommand.ExecuteNonQuery();
                        MySCommand.CommandText = "Update 任务设置 set K3= '" + 任务设置DataTable.Rows[i][7].ToString() + "' where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                        MySCommand.ExecuteNonQuery();
                        MySCommand.CommandText = "Update 任务设置 set 延时= " + (int)任务设置DataTable.Rows[i][8] + " where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                        MySCommand.ExecuteNonQuery();
                        MySCommand.CommandText = "Update 任务设置 set 截屏= '" + 任务设置DataTable.Rows[i][9].ToString() + "' where 任务名称 = '" + s1 + "' and 序号 =" + i_序号;
                        MySCommand.ExecuteNonQuery();
                    }
                    MySCommand.Transaction.Commit();
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentCell != null)
            {
                int i1 = dataGridView2.CurrentCell.RowIndex;
                if (i1 > -1)
                {
                    任务设置DataTable.Rows.RemoveAt(i1);
                    任务设置DataTable.AcceptChanges();
                    int i2 = 任务设置DataTable.Rows.Count;
                    for (int i = 0; i < i2; i++)
                    {
                        任务设置DataTable.Rows[i][0] = i + 1;
                    }
                    任务设置DataTable.AcceptChanges();
                    save任务设置(1, taskStr);
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int i1 = 任务设置DataTable.Rows.Count;
            if (i1 > 0 && CPUbool)
            {
                this.Opacity = 0;
                this.TopMost = false;
                this.Visible = false;
                int i2 = dataGridView2.CurrentRow.Index;
                string str关键词 = 任务设置DataTable.Rows[i2][1].ToString();
                string str位次 = 任务设置DataTable.Rows[i2][2].ToString();
                string str位置X = 任务设置DataTable.Rows[i2][3].ToString();
                string str位置Y = 任务设置DataTable.Rows[i2][4].ToString();
                //
                Size viewSize = new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                using (Image img = new Bitmap(viewSize.Width, viewSize.Height))
                {
                    using (Graphics gra = Graphics.FromImage(img))
                    {
                        gra.CopyFromScreen(new Point(0, 0), new Point(0, 0), img.Size);
                        //
                        //OCRResult ocr = OCREngine.DetectText(img);
                        //
                        byte[] imgTmpByte = null;
                        using (MemoryStream ms = new MemoryStream())
                        {
                            img.Save(ms, ImageFormat.Bmp);
                            imgTmpByte = ms.GetBuffer();
                        }
                        OCRResult ocr = OCREngine.DetectText(imgTmpByte);
                        if (ocr != null)
                        {
                            List<TextBlock> ltb = ocr.TextBlocks;
                            int i3 = ltb.Count;
                            if (str关键词 != "")
                            {
                                if (str位置X == "" || str位置Y == "")
                                {
                                    this.TopMost = true;
                                    MessageBox.Show("请设置正确的偏移位置。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    this.Visible = true;
                                    timer1.Enabled = true;
                                }
                                else
                                {
                                    int i_X = int.Parse(str位置X);
                                    int i_Y = int.Parse(str位置Y);
                                    int m = int.Parse(str位次);
                                    for (int x = 0; x < i3; x++)
                                    {
                                        if (ltb[x].Text.Contains(str关键词))
                                        {
                                            if (m > 0)
                                            {
                                                m--;
                                            }
                                            if (m == 0)
                                            {
                                                Pen pen = new Pen(Color.FromArgb(255, 0, 255), (float)(MyFontScale * 2));
                                                OCRPoint p1 = ltb[x].BoxPoints[0];
                                                Point p2 = new Point(p1.X + i_X, p1.Y + i_Y);
                                                if (MyDpi < 144F)
                                                {
                                                    gra.DrawImage(WhoIsRobot.Properties.Resources.from96to144, p2.X, p2.Y, 12, 19);
                                                }
                                                else if (MyDpi >= 144F && MyDpi < 192F)
                                                {
                                                    gra.DrawImage(WhoIsRobot.Properties.Resources.from144to192, p2.X, p2.Y, 18, 27);
                                                }
                                                else if (MyDpi >= 192F)
                                                {
                                                    gra.DrawImage(WhoIsRobot.Properties.Resources.Up192, p2.X, p2.Y, 25, 38);
                                                }
                                                int r_X = p1.X - smartWidth;
                                                int r_Y = p1.Y - smartHeight;
                                                gra.DrawRectangle(pen, r_X, r_Y, smartWidth * 2, smartHeight * 2);
                                                break;
                                            }
                                        }
                                    }
                                    Form f2 = new Form2();
                                    f2.BackgroundImage = img;
                                    f2.ShowDialog();
                                    this.TopMost = true;
                                    this.Visible = true;
                                    timer1.Enabled = true;
                                }
                            }
                            else
                            {
                                if (str位置X == "" || str位置Y == "")
                                {
                                    this.TopMost = true;
                                    MessageBox.Show("请设置正确的鼠标位置，F8 是获取鼠标位置的快捷键。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    this.Visible = true;
                                    timer1.Enabled = true;
                                }
                                else
                                {
                                    int i_X = int.Parse(str位置X);
                                    int i_Y = int.Parse(str位置Y);
                                    Pen pen = new Pen(Color.FromArgb(255, 0, 255), (float)(MyFontScale * 2));
                                    if (MyDpi < 144)
                                    {
                                        gra.DrawImage(WhoIsRobot.Properties.Resources.from96to144, i_X, i_Y, 12, 19);
                                    }
                                    else if (MyDpi >= 144 && MyDpi < 192)
                                    {
                                        gra.DrawImage(WhoIsRobot.Properties.Resources.from144to192, i_X, i_Y, 18, 27);
                                    }
                                    else if (MyDpi >= 192)
                                    {
                                        gra.DrawImage(WhoIsRobot.Properties.Resources.Up192, i_X, i_Y, 25, 38);
                                    }
                                    int r_X = i_X - smartWidth;
                                    int r_Y = i_Y - smartHeight;
                                    gra.DrawRectangle(pen, r_X, r_Y, smartWidth * 2, smartHeight * 2);
                                    Form f2 = new Form2();
                                    f2.BackgroundImage = img;
                                    f2.ShowDialog();
                                    this.TopMost = true;
                                    this.Visible = true;
                                    timer1.Enabled = true;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Opacity == 1)
            {
                timer1.Enabled = false;
            }
            else
            {
                Opacity += 0.1;
            }
        }

        private void dGV2moveRow(int i1, int i2, int i3)
        {
            任务设置DataTable.Rows.Clear();
            string s1 = taskStr.Replace("'", "''");
            using (SQLiteConnection MySCon = new SQLiteConnection("Data Source=" + DataFile))
            {
                MySCon.Open();
                using (SQLiteCommand MySCommand = new SQLiteCommand(MySCon))
                {
                    using (SQLiteDataAdapter MySDA = new SQLiteDataAdapter(MySCommand))
                    {
                        MySCommand.Transaction = MySCon.BeginTransaction();
                        MySCommand.CommandText = "update 任务设置 set 序号=" + i1 + " where 任务名称 = '" + s1 + "' and 序号=" + i2;
                        MySCommand.ExecuteNonQuery();
                        MySCommand.CommandText = "update 任务设置 set 序号=" + i2 + " where 任务名称 = '" + s1 + "' and 序号=" + i3;
                        MySCommand.ExecuteNonQuery();
                        MySCommand.CommandText = "update 任务设置 set 序号=" + i3 + " where 任务名称 = '" + s1 + "' and 序号=" + i1;
                        MySCommand.ExecuteNonQuery();
                        MySCommand.Transaction.Commit();
                        MySCommand.CommandText = "select 序号,关键词,位次,位置X,位置Y,K1,K2,K3,延时,截屏 from 任务设置 where 任务名称='" + s1 + "' ORDER BY 序号 ASC";
                        MySDA.Fill(任务设置DataTable);
                        dataGridView2.CurrentCell = dataGridView2[0, i3 - 1];
                        dataGridView2.CurrentCell.Selected = true;
                    }
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int i1 = 任务设置DataTable.Rows.Count;
            if (i1 > 1)
            {
                int i2 = dataGridView2.CurrentRow.Index;
                if (i2 > 0)
                {
                    dGV2moveRow(i1 + 1, i2 + 1, i2);
                }
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            int i1 = 任务设置DataTable.Rows.Count;
            if (i1 > 1)
            {
                int i2 = dataGridView2.CurrentRow.Index;
                if (i2 < i1 - 1)
                {
                    int i3 = i2 + 1;
                    dGV2moveRow(i1 + 1, i2 + 1, i2 + 2);
                }
            }
        }

        private string xmStr_org = "";
        private bool searchBL = true;

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            string s7 = textBox7.Text.Trim().ToLower();
            if (searchBL)
            {
                if (s7 != "")
                {
                    searchBL = false;
                    显示个人信息DataTable.Rows.Clear();
                    dataGridView4.DataSource = null;
                    using (BackgroundWorker xmbmBackgroundWorker = new BackgroundWorker())
                    {
                        xmbmBackgroundWorker.RunWorkerCompleted += XmbmBackgroundWorker_RunWorkerCompleted;
                        xmbmBackgroundWorker.DoWork += XmbmBackgroundWorker_DoWork;
                        xmbmBackgroundWorker.RunWorkerAsync(s7);
                    }
                }
                else
                {
                    显示个人信息DataTable.Rows.Clear();
                }
            }
        }

        private void XmbmBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string s7 = e.Argument.ToString();
            if (s7.Length == Encoding.UTF8.GetByteCount(s7))
            {
                int i1 = 个人信息DataTable.Rows.Count;
                for (int i = 0; i < i1; i++)
                {
                    if (个人信息DataTable.Rows[i][4].ToString().Contains(s7))
                    {
                        object[] obTmp = new object[4];
                        obTmp[0] = 个人信息DataTable.Rows[i][0].ToString();
                        obTmp[1] = 个人信息DataTable.Rows[i][1].ToString();
                        obTmp[2] = 个人信息DataTable.Rows[i][2].ToString();
                        obTmp[3] = 个人信息DataTable.Rows[i][3].ToString();
                        显示个人信息DataTable.Rows.Add(obTmp);
                    }
                }
            }
            else
            {
                int i1 = 个人信息DataTable.Rows.Count;
                for (int i = 0; i < i1; i++)
                {
                    if (个人信息DataTable.Rows[i][0].ToString().Contains(s7))
                    {
                        object[] obTmp = new object[4];
                        obTmp[0] = 个人信息DataTable.Rows[i][0].ToString();
                        obTmp[1] = 个人信息DataTable.Rows[i][1].ToString();
                        obTmp[2] = 个人信息DataTable.Rows[i][2].ToString();
                        obTmp[3] = 个人信息DataTable.Rows[i][3].ToString();
                        显示个人信息DataTable.Rows.Add(obTmp);
                    }
                }
            }
            xmStr_org = s7;
        }

        private void XmbmBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            string s7 = textBox7.Text.Trim().ToLower();
            if (xmStr_org != s7)
            {
                显示个人信息DataTable.Rows.Clear();
                if (s7 != "")
                {
                    using (BackgroundWorker xmbmBackgroundWorker = new BackgroundWorker())
                    {
                        xmbmBackgroundWorker.RunWorkerCompleted += XmbmBackgroundWorker_RunWorkerCompleted;
                        xmbmBackgroundWorker.DoWork += XmbmBackgroundWorker_DoWork;
                        xmbmBackgroundWorker.RunWorkerAsync(s7);
                    }
                }
            }
            else
            {
                if (s7 == "")
                {
                    显示个人信息DataTable.Rows.Clear();
                }
                dataGridView4.DataSource = dataSet2;
                dataGridView4.DataMember = 显示个人信息DataTable.TableName;
                searchBL = true;
            }
        }

        private string sfzhStr_org = "";

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            string s8 = textBox8.Text.Trim();
            if (searchBL)
            {
                if (s8.Length > 1)
                {
                    searchBL = false;
                    显示个人信息DataTable.Rows.Clear();
                    dataGridView4.DataSource = null;
                    using (BackgroundWorker sfzhBackgroundWorker = new BackgroundWorker())
                    {
                        sfzhBackgroundWorker.RunWorkerCompleted += SfzhBackgroundWorker_RunWorkerCompleted;
                        sfzhBackgroundWorker.DoWork += SfzhBackgroundWorker_DoWork;
                        sfzhBackgroundWorker.RunWorkerAsync(s8);
                    }
                }
                else
                {
                    显示个人信息DataTable.Rows.Clear();
                }
            }
        }

        private void SfzhBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string s8 = e.Argument.ToString();
            int i1 = 个人信息DataTable.Rows.Count;
            for (int i = 0; i < i1; i++)
            {
                if (个人信息DataTable.Rows[i][3].ToString().Contains(s8))
                {
                    object[] obTmp = new object[4];
                    obTmp[0] = 个人信息DataTable.Rows[i][0].ToString();
                    obTmp[1] = 个人信息DataTable.Rows[i][1].ToString();
                    obTmp[2] = 个人信息DataTable.Rows[i][2].ToString();
                    obTmp[3] = 个人信息DataTable.Rows[i][3].ToString();
                    显示个人信息DataTable.Rows.Add(obTmp);
                }
            }
            sfzhStr_org = s8;
        }

        private void SfzhBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            string s8 = textBox8.Text.Trim();
            if (sfzhStr_org != s8)
            {
                显示个人信息DataTable.Rows.Clear();
                if (s8.Length > 1)
                {
                    using (BackgroundWorker sfzhBackgroundWorker = new BackgroundWorker())
                    {
                        sfzhBackgroundWorker.RunWorkerCompleted += SfzhBackgroundWorker_RunWorkerCompleted;
                        sfzhBackgroundWorker.DoWork += SfzhBackgroundWorker_DoWork;
                        sfzhBackgroundWorker.RunWorkerAsync(s8);
                    }
                }
            }
            else
            {
                if (s8.Length < 2)
                {
                    显示个人信息DataTable.Rows.Clear();
                }
                dataGridView4.DataSource = dataSet2;
                dataGridView4.DataMember = 显示个人信息DataTable.TableName;
                searchBL = true;
            }
        }

        private void 重命名ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s1 = textBox1.Text.Trim();
            if (s1 != "")
            {
                int i1 = dataGridView1.CurrentRow.Index;
                string s2 = taskStr;
                if (s1 != s2)
                {
                    int i2 = 任务DataTable.Rows.Count;
                    ArrayList al = new ArrayList();
                    for (int i = 0; i < i2; i++)
                    {
                        if (i != i1)
                        {
                            al.Add(任务DataTable.Rows[i][0].ToString());
                        }
                    }
                    if (al.Contains(s1))
                    {
                        MessageBox.Show("相同的任务名称已存在，请修改为其他名称。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        任务DataTable.Rows[i1][0] = s1;
                        taskStr = s1;
                        s1 = s1.Replace("'", "''");
                        s2 = s2.Replace("'", "''");
                        using (SQLiteConnection MySCon = new SQLiteConnection("Data Source=" + DataFile))
                        {
                            MySCon.Open();
                            using (SQLiteCommand MySCommand = new SQLiteCommand(MySCon))
                            {
                                MySCommand.Transaction = MySCon.BeginTransaction();
                                MySCommand.CommandText = "update 任务 set 任务名称 = '" + s1 + "' where 任务名称 = '" + s2 + "'";
                                MySCommand.ExecuteNonQuery();
                                MySCommand.CommandText = "update 任务设置 set 任务名称 = '" + s1 + "' where 任务名称 = '" + s2 + "'";
                                MySCommand.ExecuteNonQuery();
                                MySCommand.Transaction.Commit();
                            }
                        }
                    }
                }
            }
        }

        private void 克隆ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s1 = taskStr;
            int i1 = 任务DataTable.Rows.Count;
            ArrayList al = new ArrayList();
            for (int i = 0; i < i1; i++)
            {
                al.Add(任务DataTable.Rows[i][0].ToString());
            }
            bool bl = true;
            int i2 = 1;
            string s2 = "";
            while (bl)
            {
                s2 = s1 + "(" + i2.ToString() + ")";
                if (!al.Contains(s2))
                {
                    bl = false;
                }
                else
                {
                    i2++;
                }
            }
            任务DataTable.Rows.Add(s2);
            save任务设置(2, s2);
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (显示个人信息DataTable.Rows.Count > 0)
            {
                int i1 = e.RowIndex;
                if (i1 > -1)
                {
                    try
                    {
                        Clipboard.Clear();
                        Clipboard.SetText(显示个人信息DataTable.Rows[i1][3].ToString());
                    }
                    catch
                    { }
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            执行任务();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            单次任务();
        }

        private void 单次任务()
        {
            if (显示个人信息DataTable.Rows.Count > 0)
            {
                if (dataGridView4.CurrentRow != null)
                {
                    int i1 = dataGridView4.CurrentRow.Index;
                    if (i1 > -1)
                    {
                        执行内容DataTable.Rows.Clear();
                        object[] ob = new object[5];
                        ob[0] = 显示个人信息DataTable.Rows[i1][0];
                        ob[1] = 显示个人信息DataTable.Rows[i1][1];
                        ob[2] = 显示个人信息DataTable.Rows[i1][2];
                        ob[3] = 显示个人信息DataTable.Rows[i1][3];
                        执行内容DataTable.Rows.Add(ob);
                        label10.Text = "1";
                        执行任务();
                    }
                }
            }
        }

        private void 执行任务()
        {
            if (CPUbool)
            {
                if (执行内容DataTable.Rows.Count > 0)
                {
                    if (任务设置DataTable.Rows.Count > 0)
                    {
                        if (checkBox1.Checked)
                        {
                            BlockInput(true);
                        }
                        Opacity = 0;
                        this.TopMost = false;
                        this.Visible = false;
                        bl中止 = false;
                        using (BackgroundWorker taskBackgroundWorker = new BackgroundWorker())
                        {
                            taskBackgroundWorker.RunWorkerCompleted += TaskBackgroundWorker_RunWorkerCompleted;
                            taskBackgroundWorker.DoWork += TaskBackgroundWorker_DoWork;
                            taskBackgroundWorker.RunWorkerAsync();
                        }
                    }
                }
            }
        }

        private int i_Log = -1;
        private string taskStr = "";

        private void TaskBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            Thread.Sleep(300);
            i_Log = -1;
            int i1 = 执行内容DataTable.Rows.Count;
            int i2 = 任务设置DataTable.Rows.Count;
            Size viewSize = new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            using (Image img = new Bitmap(viewSize.Width, viewSize.Height))
            {
                using (Graphics gra = Graphics.FromImage(img))
                {
                    ConcurrentDictionary<int, Point> pointCD = new ConcurrentDictionary<int, Point>();
                    ConcurrentDictionary<int, Rectangle> rectangleCD = new ConcurrentDictionary<int, Rectangle>();
                    for (int x = 0; x < i1; x++)
                    {
                        for (int y = 0; y < i2; y++)
                        {
                            if (bl中止)
                            {
                                i_Log = x;
                                this.Invoke(new Action(delegate
                                {
                                    this.TopMost = true;
                                    MessageBox.Show("任务已被用户中止！", "确定", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }));
                                goto AA;
                            }
                            string str关键词 = 任务设置DataTable.Rows[y][1].ToString();
                            string str位次 = 任务设置DataTable.Rows[y][2].ToString();
                            string str位置X = 任务设置DataTable.Rows[y][3].ToString();
                            string str位置Y = 任务设置DataTable.Rows[y][4].ToString();
                            string strK1 = 任务设置DataTable.Rows[y][5].ToString();
                            string strK2 = 任务设置DataTable.Rows[y][6].ToString();
                            string strK3 = 任务设置DataTable.Rows[y][7].ToString();
                            int i_延时 = (int)任务设置DataTable.Rows[y][8];
                            string str截屏 = 任务设置DataTable.Rows[y][9].ToString();
                            if (str关键词 != "" && str位置X != "" && str位置Y != "")
                            {
                                if (str截屏 == "一次")
                                {
                                    if (pointCD.TryGetValue(y, out Point p1))
                                    {
                                        CustomSendInput.移动鼠标(p1.X, p1.Y);
                                        Thread.Sleep(100);
                                    }
                                    else
                                    {
                                        try
                                        {
                                            gra.CopyFromScreen(new Point(0, 0), new Point(0, 0), img.Size);
                                        }
                                        catch
                                        {
                                            i_Log = x;
                                            this.Invoke(new Action(delegate
                                            {
                                                BlockInput(false);
                                                this.TopMost = true;
                                                MessageBox.Show("任务已被用户中止！", "确定", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            }));
                                            goto AA;
                                        }
                                        //
                                        //OCRResult ocr = OCREngine.DetectText(img);
                                        //
                                        byte[] imgTmpByte = null;
                                        using (MemoryStream ms = new MemoryStream())
                                        {
                                            img.Save(ms, ImageFormat.Bmp);
                                            imgTmpByte = ms.GetBuffer();
                                        }
                                        OCRResult ocr = OCREngine.DetectText(imgTmpByte);
                                        if (ocr == null)
                                        {
                                            i_Log = x;
                                            this.Invoke(new Action(delegate
                                            {
                                                BlockInput(false);
                                                this.TopMost = true;
                                                MessageBox.Show("第 " + (y + 1).ToString() + " 行文本解析出错，任务已退出！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            }));
                                            goto AA;
                                        }
                                        else
                                        {
                                            List<TextBlock> ltb = ocr.TextBlocks;
                                            int i3 = ltb.Count;
                                            int i_OK = 0;
                                            bool bl关键词 = true;
                                            while (i_OK < 5)
                                            {
                                                int i_位次 = int.Parse(str位次);
                                                for (int z = 0; z < i3; z++)
                                                {
                                                    if (ltb[z].Text.Contains(str关键词))
                                                    {
                                                        if (i_位次 > 0)
                                                        {
                                                            i_位次--;
                                                        }
                                                        if (i_位次 == 0)
                                                        {
                                                            int mx = ltb[z].BoxPoints[0].X + int.Parse(str位置X);
                                                            int my = ltb[z].BoxPoints[0].Y + int.Parse(str位置Y);
                                                            bl关键词 = false;
                                                            i_OK = 5;
                                                            pointCD.TryAdd(y, new Point(mx, my));
                                                            CustomSendInput.移动鼠标(mx, my);
                                                            Thread.Sleep(100);
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (bl关键词 && i_OK == 4)
                                                {
                                                    if (str关键词 == "门牌号")
                                                    {
                                                        x--;
                                                        goto mphError;
                                                    }
                                                    else
                                                    {
                                                        i_Log = x;
                                                        this.Invoke(new Action(delegate
                                                        {
                                                            BlockInput(false);
                                                            this.TopMost = true;
                                                            MessageBox.Show("没有找到第 " + (y + 1).ToString() + " 行“" + str关键词 + "”关键词！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                        }));
                                                    }
                                                    goto AA;
                                                }
                                                else if (i_OK < 5)
                                                {
                                                    Thread.Sleep(i_延时);
                                                    try
                                                    {
                                                        gra.CopyFromScreen(new Point(0, 0), new Point(0, 0), img.Size);
                                                    }
                                                    catch
                                                    {
                                                        i_Log = x;
                                                        this.Invoke(new Action(delegate
                                                        {
                                                            BlockInput(false);
                                                            this.TopMost = true;
                                                            MessageBox.Show("任务已被用户中止！", "确定", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                        }));
                                                        goto AA;
                                                    }
                                                    //
                                                    //ocr = OCREngine.DetectText(img);
                                                    //
                                                    imgTmpByte = null;
                                                    using (MemoryStream ms = new MemoryStream())
                                                    {
                                                        img.Save(ms, ImageFormat.Bmp);
                                                        imgTmpByte = ms.GetBuffer();
                                                    }
                                                    ocr = OCREngine.DetectText(imgTmpByte);
                                                    if (ocr == null)
                                                    {
                                                        i_Log = x;
                                                        this.Invoke(new Action(delegate
                                                        {
                                                            BlockInput(false);
                                                            this.TopMost = true;
                                                            MessageBox.Show("第 " + (y + 1).ToString() + " 行文本解析出错，任务已退出！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                        }));
                                                        goto AA;
                                                    }
                                                    else
                                                    {
                                                        ltb = ocr.TextBlocks;
                                                        i3 = ltb.Count;
                                                        i_OK++;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else if (str截屏 == "每次")
                                {
                                    bool r_bl = true;
                                    if (rectangleCD.TryGetValue(y, out Rectangle r1) && pointCD.TryGetValue(y, out Point p1))
                                    {
                                        using (Image imgTmp = new Bitmap(r1.Size.Width, r1.Size.Height))
                                        {
                                            using (Graphics graTmp = Graphics.FromImage(imgTmp))
                                            {
                                                try
                                                {
                                                    graTmp.CopyFromScreen(r1.Location, new Point(0, 0), r1.Size);
                                                }
                                                catch
                                                {
                                                    i_Log = x;
                                                    this.Invoke(new Action(delegate
                                                    {
                                                        BlockInput(false);
                                                        this.TopMost = true;
                                                        MessageBox.Show("任务已被用户中止！", "确定", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                    }));
                                                    goto AA;
                                                }
                                                //
                                                //OCRResult ocr = OCREngine.DetectText(imgTmp);
                                                //
                                                byte[] imgTmpByte = null;
                                                using (MemoryStream ms = new MemoryStream())
                                                {
                                                    imgTmp.Save(ms, ImageFormat.Bmp);
                                                    imgTmpByte = ms.GetBuffer();
                                                }
                                                OCRResult ocr = OCREngine.DetectText(imgTmpByte);
                                                if (ocr != null)
                                                {
                                                    List<TextBlock> ltb = ocr.TextBlocks;
                                                    int i3 = ltb.Count;
                                                    for (int z = 0; z < i3; z++)
                                                    {
                                                        if (ltb[z].Text.Contains(str关键词))
                                                        {
                                                            int mx = r1.X + ltb[z].BoxPoints[0].X + int.Parse(str位置X);
                                                            int my = r1.Y + ltb[z].BoxPoints[0].Y + int.Parse(str位置Y);
                                                            if (Math.Abs(mx - p1.X) < 160 && Math.Abs(my - p1.Y) < 90)
                                                            {
                                                                r_bl = false;
                                                                CustomSendInput.移动鼠标(mx, my);
                                                                Thread.Sleep(100);
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (r_bl)
                                    {
                                        try
                                        {
                                            gra.CopyFromScreen(new Point(0, 0), new Point(0, 0), img.Size);
                                        }
                                        catch
                                        {
                                            i_Log = x;
                                            this.Invoke(new Action(delegate
                                            {
                                                BlockInput(false);
                                                this.TopMost = true;
                                                MessageBox.Show("任务已被用户中止！", "确定", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            }));
                                            goto AA;
                                        }
                                        //
                                        //OCRResult ocr = OCREngine.DetectText(img);
                                        //
                                        byte[] imgTmpByte = null;
                                        using (MemoryStream ms = new MemoryStream())
                                        {
                                            img.Save(ms, ImageFormat.Bmp);
                                            imgTmpByte = ms.GetBuffer();
                                        }
                                        OCRResult ocr = OCREngine.DetectText(imgTmpByte);
                                        if (ocr == null)
                                        {
                                            i_Log = x;
                                            this.Invoke(new Action(delegate
                                            {
                                                BlockInput(false);
                                                this.TopMost = true;
                                                MessageBox.Show("第 " + (y + 1).ToString() + " 行文本解析出错，任务已退出！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            }));
                                            goto AA;
                                        }
                                        else
                                        {
                                            List<TextBlock> ltb = ocr.TextBlocks;
                                            int i3 = ltb.Count;
                                            int i_OK = 0;
                                            bool bl关键词 = true;
                                            while (i_OK < 5)
                                            {
                                                int i_位次 = int.Parse(str位次);
                                                for (int z = 0; z < i3; z++)
                                                {
                                                    if (ltb[z].Text.Contains(str关键词))
                                                    {
                                                        if (i_位次 > 0)
                                                        {
                                                            i_位次--;
                                                        }
                                                        if (i_位次 == 0)
                                                        {
                                                            int l_X = ltb[z].BoxPoints[0].X;
                                                            int smart_X_Top = l_X - smartWidth;
                                                            if (smart_X_Top < 0)
                                                            {
                                                                smart_X_Top = 0;
                                                            }
                                                            int l_Y = ltb[z].BoxPoints[0].Y;
                                                            int smart_Y_Top = l_Y - smartHeight;
                                                            if (smart_Y_Top < 0)
                                                            {
                                                                smart_Y_Top = 0;
                                                            }
                                                            int smart_X_Bottom = l_X + smartWidth;
                                                            if (smart_X_Bottom > viewSize.Width)
                                                            {
                                                                smart_X_Bottom = smartWidth * 2 + viewSize.Width - smart_X_Bottom;
                                                            }
                                                            else
                                                            {
                                                                smart_X_Bottom = smartWidth * 2;
                                                            }
                                                            int smart_Y_Bottom = l_Y + smartHeight;
                                                            if (smart_Y_Bottom > viewSize.Height)
                                                            {
                                                                smart_Y_Bottom = smartHeight * 2 + viewSize.Height - smart_Y_Bottom;
                                                            }
                                                            else
                                                            {
                                                                smart_Y_Bottom = smartHeight * 2;
                                                            }
                                                            bl关键词 = false;
                                                            i_OK = 5;
                                                            rectangleCD.TryRemove(y, out Rectangle r2);
                                                            rectangleCD.TryAdd(y, new Rectangle(smart_X_Top, smart_Y_Top, smart_X_Bottom, smart_Y_Bottom));
                                                            int mx = l_X + int.Parse(str位置X);
                                                            int my = l_Y + int.Parse(str位置Y);
                                                            pointCD.TryRemove(y, out Point p2);
                                                            pointCD.TryAdd(y, new Point(mx, my));
                                                            CustomSendInput.移动鼠标(mx, my);
                                                            Thread.Sleep(100);
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (bl关键词 && i_OK == 4)
                                                {
                                                    if (str关键词 == "门牌号")
                                                    {
                                                        x--;
                                                        goto mphError;
                                                    }
                                                    else
                                                    {
                                                        i_Log = x;
                                                        this.Invoke(new Action(delegate
                                                        {
                                                            BlockInput(false);
                                                            this.TopMost = true;
                                                            MessageBox.Show("没有找到第 " + (y + 1).ToString() + " 行“" + str关键词 + "”关键词！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                        }));
                                                    }
                                                    goto AA;
                                                }
                                                else if (i_OK < 5)
                                                {
                                                    Thread.Sleep(i_延时);
                                                    try
                                                    {
                                                        gra.CopyFromScreen(new Point(0, 0), new Point(0, 0), img.Size);
                                                    }
                                                    catch
                                                    {
                                                        i_Log = x;
                                                        this.Invoke(new Action(delegate
                                                        {
                                                            BlockInput(false);
                                                            this.TopMost = true;
                                                            MessageBox.Show("任务已被用户中止！", "确定", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                        }));
                                                        goto AA;
                                                    }
                                                    //
                                                    //ocr = OCREngine.DetectText(img);
                                                    //
                                                    imgTmpByte = null;
                                                    using (MemoryStream ms = new MemoryStream())
                                                    {
                                                        img.Save(ms, ImageFormat.Bmp);
                                                        imgTmpByte = ms.GetBuffer();
                                                    }
                                                    ocr = OCREngine.DetectText(imgTmpByte);
                                                    if (ocr == null)
                                                    {
                                                        i_Log = x;
                                                        this.Invoke(new Action(delegate
                                                        {
                                                            BlockInput(false);
                                                            this.TopMost = true;
                                                            MessageBox.Show("第 " + (y + 1).ToString() + " 行文本解析出错，任务已退出！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                        }));
                                                        goto AA;
                                                    }
                                                    else
                                                    {
                                                        ltb = ocr.TextBlocks;
                                                        i3 = ltb.Count;
                                                        i_OK++;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else if (str位置X != "" && str位置Y != "")
                            {
                                CustomSendInput.移动鼠标(int.Parse(str位置X), int.Parse(str位置Y));
                                Thread.Sleep(100);
                            }
                            if (strK1 != "(空)")
                            {
                                switch (strK1)
                                {
                                    case "Tab":
                                        CustomSendInput.Tab键();
                                        break;
                                    case "Delete":
                                        CustomSendInput.删除键();
                                        break;
                                    case "Backspace":
                                        CustomSendInput.退格键();
                                        break;
                                    case "Enter":
                                        CustomSendInput.回车();
                                        break;
                                    case "Home":
                                        CustomSendInput.Home键();
                                        break;
                                    case "End":
                                        CustomSendInput.End键();
                                        break;
                                    case "向上":
                                        CustomSendInput.向上();
                                        break;
                                    case "向下":
                                        CustomSendInput.向下();
                                        break;
                                    case "向左":
                                        CustomSendInput.向左();
                                        break;
                                    case "向右":
                                        CustomSendInput.向右();
                                        break;
                                    case "鼠标右键":
                                        CustomSendInput.鼠标右键();
                                        break;
                                    case "鼠标单击":
                                        CustomSendInput.鼠标单击();
                                        break;
                                    case "鼠标双击":
                                        CustomSendInput.鼠标双击();
                                        break;
                                }
                                Thread.Sleep(i_延时);
                            }
                            if (strK2 != "(空)")
                            {
                                switch (strK2)
                                {
                                    case "Enter":
                                        CustomSendInput.回车();
                                        break;
                                    case "Ctrl+A":
                                        CustomSendInput.全选();
                                        break;
                                    case "Ctrl+C":
                                        CustomSendInput.复制();
                                        break;
                                    case "Ctrl+V":
                                        CustomSendInput.粘贴();
                                        break;
                                    case "PageUp":
                                        CustomSendInput.向上翻页();
                                        break;
                                    case "PageDown":
                                        CustomSendInput.向下翻页();
                                        break;
                                    case "鼠标单击":
                                        CustomSendInput.鼠标单击();
                                        break;
                                    case "鼠标双击":
                                        CustomSendInput.鼠标双击();
                                        break;
                                    case "复制 1":
                                        this.Invoke(new Action(delegate
                                        {
                                            Clipboard.Clear();
                                            string strTmp = 执行内容DataTable.Rows[x][0].ToString();
                                            if (strTmp != "")
                                            {
                                                Clipboard.SetText(strTmp);
                                            }
                                        }));
                                        break;
                                    case "复制 2":
                                        this.Invoke(new Action(delegate
                                        {
                                            Clipboard.Clear();
                                            string strTmp = 执行内容DataTable.Rows[x][1].ToString();
                                            if (strTmp != "")
                                            {
                                                Clipboard.SetText(strTmp);
                                            }
                                        }));
                                        break;
                                    case "复制 3":
                                        this.Invoke(new Action(delegate
                                        {
                                            Clipboard.Clear();
                                            string strTmp = 执行内容DataTable.Rows[x][2].ToString();
                                            if (strTmp != "")
                                            {
                                                Clipboard.SetText(strTmp);
                                            }
                                        }));
                                        break;
                                    case "复制 4":
                                        this.Invoke(new Action(delegate
                                        {
                                            Clipboard.Clear();
                                            string strTmp = 执行内容DataTable.Rows[x][3].ToString();
                                            if (strTmp != "")
                                            {
                                                Clipboard.SetText(strTmp);
                                            }
                                        }));
                                        break;
                                    case "复制 5":
                                        this.Invoke(new Action(delegate
                                        {
                                            Clipboard.Clear();
                                            string strTmp = 执行内容DataTable.Rows[x][4].ToString();
                                            if (strTmp != "")
                                            {
                                                Clipboard.SetText(strTmp);
                                            }
                                        }));
                                        break;
                                }
                                Thread.Sleep(i_延时);
                            }
                            if (strK3 != "(空)")
                            {
                                switch (strK3)
                                {
                                    case "Enter":
                                        CustomSendInput.回车();
                                        break;
                                    case "Ctrl+A":
                                        CustomSendInput.全选();
                                        break;
                                    case "Ctrl+C":
                                        CustomSendInput.复制();
                                        break;
                                    case "Ctrl+V":
                                        CustomSendInput.粘贴();
                                        break;
                                    case "PageUp":
                                        CustomSendInput.向上翻页();
                                        break;
                                    case "PageDown":
                                        CustomSendInput.向下翻页();
                                        break;
                                    case "鼠标单击":
                                        CustomSendInput.鼠标单击();
                                        break;
                                    case "鼠标双击":
                                        CustomSendInput.鼠标双击();
                                        break;
                                    case "复制 1":
                                        this.Invoke(new Action(delegate
                                        {
                                            Clipboard.Clear();
                                            string strTmp = 执行内容DataTable.Rows[x][0].ToString();
                                            if (strTmp != "")
                                            {
                                                Clipboard.SetText(strTmp);
                                            }
                                        }));
                                        break;
                                    case "复制 2":
                                        this.Invoke(new Action(delegate
                                        {
                                            Clipboard.Clear();
                                            string strTmp = 执行内容DataTable.Rows[x][1].ToString();
                                            if (strTmp != "")
                                            {
                                                Clipboard.SetText(strTmp);
                                            }
                                        }));
                                        break;
                                    case "复制 3":
                                        this.Invoke(new Action(delegate
                                        {
                                            Clipboard.Clear();
                                            string strTmp = 执行内容DataTable.Rows[x][2].ToString();
                                            if (strTmp != "")
                                            {
                                                Clipboard.SetText(strTmp);
                                            }
                                        }));
                                        break;
                                    case "复制 4":
                                        this.Invoke(new Action(delegate
                                        {
                                            Clipboard.Clear();
                                            string strTmp = 执行内容DataTable.Rows[x][3].ToString();
                                            if (strTmp != "")
                                            {
                                                Clipboard.SetText(strTmp);
                                            }
                                        }));
                                        break;
                                    case "复制 5":
                                        this.Invoke(new Action(delegate
                                        {
                                            Clipboard.Clear();
                                            string strTmp = 执行内容DataTable.Rows[x][4].ToString();
                                            if (strTmp != "")
                                            {
                                                Clipboard.SetText(strTmp);
                                            }
                                        }));
                                        break;
                                }
                                Thread.Sleep(i_延时);
                            }
                        mphError: string s1 = "";
                        }
                    }
                }
            }
        AA: string s2 = "";
        }

        private void TaskBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BlockInput(false);
            this.TopMost = true;
            if (i_Log > -1)
            {
                if (i_Log < (执行内容DataTable.Rows.Count - 1))
                {
                    i_Log++;
                    StringBuilder sb = new StringBuilder();
                    sb.Append("即将执行第 " + i_Log.ToString() + " 行任务\r\n\r\n");
                    for (int z = 0; z < 执行内容DataTable.Columns.Count; z++)
                    {
                        if (执行内容DataTable.Rows[i_Log][z].ToString() != "")
                        {
                            sb.Append(执行内容DataTable.Rows[i_Log][z].ToString() + "\r\n");
                        }
                    }
                    using (StreamWriter sw = new StreamWriter(taskLogFile, false))
                    {
                        sw.WriteLine(sb.ToString().TrimEnd());
                    }
                }
            }
            else
            {
                File.Delete(taskLogFile);
                MessageBox.Show("任务执行完成。", "确定", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            this.Visible = true;
            Opacity = 1;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterHotKey(this.Handle, 1578);
            UnregisterHotKey(this.Handle, 1579);
        }

        private bool bl中止 = false;

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312)
            {
                if (m.WParam.ToInt32() == 1578)
                {
                    IntPtr r = CustomSendInput.GetCursorPos(out Point p);
                    if (r.ToInt32() == 1 && this.Visible)
                    {
                        textBox2.Text = p.X.ToString();
                        textBox3.Text = p.Y.ToString();
                    }
                }
                else if (m.WParam.ToInt32() == 1579)
                {
                    if (this.Visible)
                    {
                        int i1 = tabControl1.SelectedIndex;
                        if (i1 == 0)
                        {
                            执行任务();
                        }
                        else if (i1 == 1)
                        {
                            单次任务();
                        }
                    }
                    else
                    {
                        bl中止 = true;
                    }
                }
            }
            base.WndProc(ref m);
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (任务设置DataTable.Rows.Count > 0)
            {
                int i2 = e.RowIndex;
                if (i2 > -1)
                {
                    using (SQLiteConnection MySCon = new SQLiteConnection("Data Source=" + DataFile))
                    {
                        MySCon.Open();
                        using (SQLiteCommand MySCommand = new SQLiteCommand(MySCon))
                        {
                            int i_序号 = i2 + 1;
                            int i3 = e.ColumnIndex;
                            string s1 = 任务设置DataTable.Rows[i2][i3].ToString();
                            if (i3 == 1)
                            {
                                MySCommand.CommandText = "update 任务设置 set 关键词 ='" + s1.Replace("'", "''") + "' where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                            }
                            else if (i3 == 2)
                            {
                                uint u_位次 = 1;
                                try
                                {
                                    u_位次 = uint.Parse(s1);
                                }
                                catch
                                { }
                                if (u_位次 == 0)
                                {
                                    u_位次 = 1;
                                }
                                任务设置DataTable.Rows[i2][i3] = u_位次;
                                MySCommand.CommandText = "update 任务设置 set 位次 =" + u_位次 + " where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                            }
                            else if (i3 == 3)
                            {
                                if (s1 == "")
                                {
                                    任务设置DataTable.Rows[i2][i3] = DBNull.Value;
                                    MySCommand.CommandText = "update 任务设置 set 位置X = null where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                                }
                                else
                                {
                                    int iX = int.Parse(s1);
                                    任务设置DataTable.Rows[i2][i3] = iX;
                                    MySCommand.CommandText = "update 任务设置 set 位置X =" + iX + " where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                                }
                            }
                            else if (i3 == 4)
                            {
                                if (s1 == "")
                                {
                                    任务设置DataTable.Rows[i2][i3] = DBNull.Value;
                                    MySCommand.CommandText = "update 任务设置 set 位置Y = null where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                                }
                                else
                                {
                                    int iY = int.Parse(s1);
                                    任务设置DataTable.Rows[i2][i3] = iY;
                                    MySCommand.CommandText = "update 任务设置 set 位置Y =" + iY + " where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                                }
                            }
                            else if (i3 == 5)
                            {
                                MySCommand.CommandText = "update 任务设置 set K1 ='" + s1 + "' where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                            }
                            else if (i3 == 6)
                            {
                                MySCommand.CommandText = "update 任务设置 set K2 ='" + s1 + "' where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                            }
                            else if (i3 == 7)
                            {
                                MySCommand.CommandText = "update 任务设置 set K3 ='" + s1 + "' where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                            }
                            else if (i3 == 8)
                            {
                                uint u_延时 = 200;
                                try
                                {
                                    u_延时 = uint.Parse(s1);
                                    if (u_延时 < 10)
                                    {
                                        u_延时 = 10;
                                    }
                                }
                                catch
                                { }
                                任务设置DataTable.Rows[i2][i3] = u_延时;
                                MySCommand.CommandText = "update 任务设置 set 延时 =" + u_延时 + " where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                            }
                            else if (i3 == 9)
                            {
                                MySCommand.CommandText = "update 任务设置 set 截屏 ='" + s1 + "' where 任务名称 = '" + taskStr.Replace("'", "''") + "' and 序号 =" + i_序号;
                            }
                            MySCommand.ExecuteNonQuery();
                        }
                    }
                }
            }
        }

        private void 查看日志ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (File.Exists(taskLogFile))
            {
                Process.Start("notepad.exe", taskLogFile);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                tabControl1.Enabled = false;
                using (BackgroundWorker loadgrxxBackgroundWorker = new BackgroundWorker())
                {
                    loadgrxxBackgroundWorker.RunWorkerCompleted += LoadgrxxBackgroundWorker_RunWorkerCompleted;
                    loadgrxxBackgroundWorker.DoWork += LoadgrxxBackgroundWorker_DoWork;
                    loadgrxxBackgroundWorker.RunWorkerAsync(openFileDialog2.FileName);
                }
            }
        }

        private void LoadgrxxBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                using (FileStream FS = File.OpenRead(e.Argument.ToString()))
                {
                    using (BinaryReader BR = new BinaryReader(FS))
                    {
                        IWorkbook IWB;
                        if (string.Compare(Encoding.ASCII.GetString(BR.ReadBytes(2)), "PK", StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            FS.Seek(0, SeekOrigin.Begin);
                            IWB = new XSSFWorkbook(FS);
                        }
                        else
                        {
                            FS.Seek(0, SeekOrigin.Begin);
                            IWB = new HSSFWorkbook(FS);
                        }
                        ISheet MySheet = IWB.GetSheetAt(0);
                        if (MySheet != null)
                        {
                            int MaxIndex = MySheet.LastRowNum;
                            if (MaxIndex > 0)
                            {
                                IRow MyRow = MySheet.GetRow(0);
                                int i1 = MyRow.Cells.Count;
                                int xm = -1, xb = -1, nl = -1, sfzh = -1;
                                int i2 = 0;
                                try
                                {
                                    for (int i = 0; i < i1; i++)
                                    {
                                        switch (MyRow.Cells[i].StringCellValue.Trim())
                                        {
                                            case "姓名":
                                                xm = i;
                                                i2++;
                                                break;
                                            case "性别":
                                                xb = i;
                                                i2++;
                                                break;
                                            case "年龄":
                                                nl = i;
                                                i2++;
                                                break;
                                            case "身份证号":
                                                sfzh = i;
                                                i2++;
                                                break;
                                        }
                                    }
                                }
                                catch
                                { }
                                if (i2 == 4)
                                {
                                    for (int i = 1; i <= MaxIndex; i++)
                                    {
                                        object[] obTmp = new object[5];
                                        string xmStr = "";
                                        MyRow = MySheet.GetRow(i);
                                        i1 = MyRow.Cells.Count;
                                        for (int y = 0; y < i1; y++)
                                        {
                                            if (y == xm)
                                            {
                                                if (MyRow.GetCell(y) != null)
                                                {
                                                    if (MyRow.GetCell(y).CellType == CellType.Numeric)
                                                    {
                                                        xmStr = MyRow.GetCell(y).NumericCellValue.ToString();
                                                    }
                                                    else
                                                    {
                                                        xmStr = MyRow.GetCell(y).StringCellValue.Trim();
                                                    }
                                                }
                                            }
                                            else if (y == xb)
                                            {
                                                if (MyRow.GetCell(y) != null)
                                                {
                                                    if (MyRow.GetCell(y).CellType == CellType.Numeric)
                                                    {
                                                        obTmp[1] = MyRow.GetCell(y).NumericCellValue.ToString();
                                                    }
                                                    else
                                                    {
                                                        obTmp[1] = MyRow.GetCell(y).StringCellValue.Trim();
                                                    }
                                                }
                                            }
                                            else if (y == nl)
                                            {
                                                if (MyRow.GetCell(y) != null)
                                                {
                                                    if (MyRow.GetCell(y).CellType == CellType.Numeric)
                                                    {
                                                        obTmp[2] = MyRow.GetCell(y).NumericCellValue.ToString();
                                                    }
                                                    else
                                                    {
                                                        obTmp[2] = MyRow.GetCell(y).StringCellValue.Trim();
                                                    }
                                                }
                                            }
                                            else if (y == sfzh)
                                            {
                                                if (MyRow.GetCell(y) != null)
                                                {
                                                    if (MyRow.GetCell(y).CellType == CellType.Numeric)
                                                    {
                                                        obTmp[3] = MyRow.GetCell(y).NumericCellValue.ToString();
                                                    }
                                                    else
                                                    {
                                                        obTmp[3] = MyRow.GetCell(y).StringCellValue.Trim();
                                                    }
                                                }
                                            }
                                        }
                                        if (xmStr != "")
                                        {
                                            obTmp[0] = xmStr;
                                            obTmp[4] = GetAllZM(xmStr);
                                            个人信息DataTable.Rows.Add(obTmp);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.Invoke(new Action(delegate
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
            }
        }

        private void LoadgrxxBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            tabControl1.Enabled = true;
        }

        private void dataGridView2_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                MessageBox.Show(e.Exception.Message + "\r\n位次应为大于零的正整数。");
            }
            else if (e.ColumnIndex == 2)
            {
                MessageBox.Show(e.Exception.Message + "\r\n位置(X)应为整数。");
            }
            else if (e.ColumnIndex == 3)
            {
                MessageBox.Show(e.Exception.Message + "\r\n位置(Y)应为整数。");
            }
            else if (e.ColumnIndex == 7)
            {
                MessageBox.Show(e.Exception.Message + "\r\n延时应为不小于 10 的正整数。");
            }
        }
    }
}
