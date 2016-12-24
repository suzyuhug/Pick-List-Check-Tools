


using System.Drawing;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;

using System.Data;

using System.IO;

using System.Threading;
using System.Windows.Forms;
using System.Drawing.Printing;

namespace Pick_List_Check_Tools
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
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
                dt.Columns.Add("Status", typeof(string));
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



                                    

                                  
                                    sheet.GetRow(1).CreateCell(14).SetCellValue("电源：");
                                    sheet.GetRow(1).CreateCell(15).SetCellValue("无");
                                    sheet.GetRow(2).CreateCell(14).SetCellValue("DSP：");
                                    sheet.GetRow(2).CreateCell(15).SetCellValue("无");
                                    sheet.GetRow(3).CreateCell(14).SetCellValue("散热片：");
                                    sheet.GetRow(3).CreateCell(15).SetCellValue("无");


                                    FileStream file = new FileStream(filePath, FileMode.Create);
                            workbook.Write(file);
                            file.Close();







                            // MessageBox.Show($"{itemid.ToString()}  {compoentid.ToString()}  {eqtyid.ToString()}  {categoryid.ToString()}    {mccheckid.ToString()}");

                            //======================================================================================================================


                            PB.Maximum = rowCount * cellCount-2;

                                    for (int i = Header + 1; i <= rowCount; ++i)
                                    {
                                        row = sheet.GetRow(i);
                                        if (row == null) continue;

                                        // dataRow = dataTable.NewRow();
                                        string itemstr = null, compoentstr = null, eqtystr = null, categorystr = null;
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
                                                    itemstr = strvalue;
                                                }

                                                if (j == compoentid)
                                                {
                                                    compoentstr = strvalue;

                                                }
                                                if (j == eqtyid)
                                                {
                                                    eqtystr = strvalue;
                                                }
                                                if (j == categoryid)
                                                {
                                                    categorystr = strvalue;
                                                }
                                            }
                                        }
                                        bool bl = false; string dtStatus = null;
                                        bool dspbool = false, powerbool = false, hebool = false;
                                        if (itemstr != null && compoentstr != null && eqtystr != null && categorystr != null)
                                        {
                                            for (int n = 0; n < dt.Rows.Count; n++)
                                            {
                                                string dtitem = dt.Rows[n]["Item Number"].ToString();
                                                string dtcomt = dt.Rows[n]["Component"].ToString();
                                                string dtcate = dt.Rows[n]["Category"].ToString();
                                                string dteqty = dt.Rows[n]["E Qty"].ToString();
                                                dtStatus = dt.Rows[n]["Status"].ToString();
                                                if (dtitem == itemstr && dtcomt == compoentstr && dtcate == categorystr && dteqty == eqtystr)
                                                {
                                                    if (compoentstr == "TDN-601-979-18") powerbool = true;
                                                    if (compoentstr == "TDN-873-572-21") hebool = true;
                                                    if (compoentstr == "TDN-624-103-00") dspbool = true;
                                                    bl = true;
                                                    break;
                                                }



                                            }

                                        }


                                        if (bl)
                                        {
                                            loglabel.Text = $"正在写入：{dtStatus}";

                                            Application.DoEvents();
                                            sheet.GetRow(i).GetCell(mccheckid).SetCellValue(dtStatus);
                                            //row.CreateCell(mccheckid).SetCellValue("EM");

                                            if (powerbool)
                                            {
                                                sheet.GetRow(1).CreateCell(14).SetCellValue("电源：");
                                                sheet.GetRow(1).CreateCell(15).SetCellValue("有");
                                               
                                            }
                                            if (dspbool)
                                            {
                                                
                                                sheet.GetRow(2).CreateCell(14).SetCellValue("DSP：");
                                                sheet.GetRow(2).CreateCell(15).SetCellValue("有");
                                               


                                            }
                                            if (hebool)
                                            {
                                               
                                                sheet.GetRow(3).CreateCell(14).SetCellValue("散热片：");
                                                sheet.GetRow(3).CreateCell(15).SetCellValue("有");


                                            }
                                            FileStream file1 = new FileStream(filePath, FileMode.Create);

                                            workbook.Write(file1);
                                            file1.Close();


                                        }


                                        //dataTable.Rows.Add(dataRow);
                                    }

                                    //======================================打印文件==============================================

                                    PrintPriviewExcelFile(filePath);



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
            UpdateClass.UpdateFrom("PickList-Tool");
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


        public void PrintPriviewExcelFile(string filePath)
        {
            Microsoft.Office.Interop.Excel.Application myexcel = new Microsoft.Office.Interop.Excel.Application();
            myexcel.Visible = false;
            myexcel.Application.DisplayAlerts = false;
            myexcel.Workbooks.Open(filePath);
            Microsoft.Office.Interop.Excel.Worksheet mysheet = (Microsoft.Office.Interop.Excel.Worksheet)myexcel.Worksheets[1];
            mysheet.Activate();


            myexcel.ActiveSheet.PrintOut();
            myexcel.Workbooks.Close();
            myexcel.Quit();

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
          
        }
    }
}

        