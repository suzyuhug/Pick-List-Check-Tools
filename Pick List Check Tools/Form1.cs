
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Pick_List_Check_Tools
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        private void button1_Click(object sender, EventArgs e)
        {

          


            string filepath = @"d:\456.xls";
            ExcelToDataTable(filepath, false);


        }
        private void ExcelToDataTable(string filePath, bool isColumnName)
        {

            FileStream fs = null;

            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;

            try
            {
                PB.Value = 0;
                //===================================获取本地数据===================================================
                DataTable dt = new DataTable();
                dt.Columns.Add("Item Number", typeof(string));
                dt.Columns.Add("Component", typeof(string));
                dt.Columns.Add("Category", typeof(string));
                dt.Columns.Add("E Qty", typeof(string));
                dt = CsvHelper.csv2dt($"{Application.StartupPath.ToString()}\\csv.csv", 1, dt);

                //=================================================================================================

                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);

                    // 2003版本
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet                      
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数
                           // loglabel.Text = rowCount.ToString();//-------------------------------
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行
                                int cellCount = firstRow.LastCellNum;//列数
                                //MessageBox.Show(cellCount.ToString());

                                row = sheet.GetRow(0);
                                cell = row.GetCell(0);
                                string str = null;
                                if (cell == null)
                                {
                                    str = "";
                                }
                                else
                                {
                                    str = cell.StringCellValue.ToString();
                                }


                                // loglabel.Text = cell.StringCellValue;
                                if (str == "Teradyne Operation Pick List")
                                {
                                    //============================================================
                                    //=============读取PO#=======================================
                                    row = sheet.GetRow(1);
                                    cell = row.GetCell(0);
                                    loglabel.Text = "正在读取PO文件";
                                    Application.DoEvents();
                                    if (cell != null) PONumLabel.Text = cell.StringCellValue.ToString();
                                    //=========================================================



                                    //==================判断表头指定行==========================
                                    int Header = 0;//表头的行号
                                    int itemid = -1, compoentid = -1, eqtyid = -1, categoryid = -1, mccheckid = -1;
                                    for (int i = 4; i < 10; i++)
                                    {
                                        row = sheet.GetRow(i);
                                        if (row == null) continue;
                                        cell = row.GetCell(0);

                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue.ToString() == "Line")
                                            {
                                                firstRow = sheet.GetRow(i);//第一行
                                                cellCount = firstRow.LastCellNum;//列数
                                                Header = i;
                                                for (int j = row.FirstCellNum; j < cellCount; j++)
                                                {

                                                    cell = row.GetCell(j);
                                                    if (cell != null)
                                                    {
                                                        loglabel.Text = $"查找表头：{ cell.StringCellValue.ToString() }";
                                                        Application.DoEvents();
                                                        // MessageBox.Show(cell.StringCellValue.ToString());
                                                        if (cell.StringCellValue.ToString() == "Item Number") itemid = j;
                                                        if (cell.StringCellValue.ToString() == "Component") compoentid = j;
                                                        if (cell.StringCellValue.ToString() == "E Qty") eqtyid = j;
                                                        if (cell.StringCellValue.ToString() == "Category") categoryid = j;
                                                        if (cell.StringCellValue.ToString() == "MC Check") mccheckid = j;


                                                    }
                                                }
                                            }
                                        }
                                    }





                                    // MessageBox.Show($"{itemid.ToString()}  {compoentid.ToString()}  {eqtyid.ToString()}  {categoryid.ToString()}    {mccheckid.ToString()}");

                                    //======================================================================================================================


                                    PB.Maximum = rowCount * cellCount-2;

                                    for (int i = Header + 1; i <= rowCount; ++i)
                                    {
                                        row = sheet.GetRow(i);
                                        if (row == null) continue;

                                        // dataRow = dataTable.NewRow();
                                        bool itembool = false, compoentbool = false, eqtybool = false, categorybool = false;
                                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                                        {
                                            PB.Value = PB.Value + 1;
                                            cell = row.GetCell(j);
                                            string strvalue = "";
                                            if (cell != null)
                                            {

                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank:
                                                        strvalue = "";
                                                        break;
                                                    case CellType.Numeric:
                                                        short format = cell.CellStyle.DataFormat;

                                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                                            strvalue = cell.DateCellValue.ToString();
                                                        else
                                                            strvalue = cell.NumericCellValue.ToString();
                                                        break;
                                                    case CellType.String:
                                                        strvalue = cell.StringCellValue;
                                                        break;
                                                }
                                                //DataRow[] ExistpoArr;
                                                loglabel.Text = $"正在对比数据：{strvalue}";
                                                Application.DoEvents();
                                                if (j == itemid)
                                                {
                                                    DataRow[] ExistpoArr = dt.Select("[Item Number]='" + strvalue + "'");
                                                    if (ExistpoArr.Length > 0)
                                                    {
                                                        itembool = true;
                                                    }
                                                }

                                                if (j == compoentid)
                                                {
                                                    DataRow[] ExistpoArr = dt.Select("Component='" + strvalue + "'");
                                                    if (ExistpoArr.Length > 0)
                                                    {

                                                        compoentbool = true;
                                                    }
                                                }
                                                if (j == eqtyid)
                                                {
                                                    DataRow[] ExistpoArr = dt.Select("[E Qty]='" + strvalue + "'");
                                                    if (ExistpoArr.Length > 0)
                                                    {
                                                        eqtybool = true;
                                                    }
                                                }
                                                if (j == categoryid)
                                                {
                                                    DataRow[] ExistpoArr = dt.Select("Category='" + strvalue + "'");
                                                    if (ExistpoArr.Length > 0)
                                                    {
                                                        categorybool = true;
                                                    }
                                                }
                                            }
                                        }

                                        if (itembool && compoentbool && eqtybool && categorybool)
                                        {
                                            loglabel.Text = $"正在写入：EM";
                                            Application.DoEvents();
                                            sheet.GetRow(i).GetCell(mccheckid).SetCellValue("东方不败");
                                            //row.CreateCell(mccheckid).SetCellValue("EM");
                                            FileStream file = new FileStream(filePath, FileMode.Create);

                                            workbook.Write(file);
                                            file.Close();
                                        }

                                        //dataTable.Rows.Add(dataRow);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show(" 此文件不是Pick List文件，无法校验！", "PO校验工具", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                    }
                }
                PB.Maximum = 1;
                PB.Value = 1;
                loglabel.Text = "对比完成";
                this.Close();
                this.Dispose(); 
                Application.Exit();

            }
            catch (Exception)
            {
                if (fs != null)
                {
                    fs.Close();

                    MessageBox.Show("PO文件校验错误，原因不详！", "PO校验工具", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                    this.Dispose();
                    Application.Exit();
                }
                // return null;
            }
        }
        Thread thread;
        private void Form1_Load(object sender, EventArgs e)
        {
            int ScreenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int ScreenHeight = Screen.PrimaryScreen.WorkingArea.Height;
            int x = ScreenWidth - this.Width - 5;
            int y = ScreenHeight - this.Height - 5;
            this.Location = new Point(x, y);
            thread = new Thread(loadupdate);
            thread.Start();

        }
        private void loadupdate()
        {

           
            string filepath = Program.str;
            this.Invoke((EventHandler)delegate
            {

                ExcelToDataTable(filepath, false);
            });
            thread.Abort();

        }
    }
}

        