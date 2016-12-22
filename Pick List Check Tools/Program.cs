using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Pick_List_Check_Tools
{
    static class Program
    {

        public static string str = null;
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            if (args.Length != 0)
            {
                str = args[0].ToString();
                //MessageBox.Show(str);
               // str = @"D:\456.xls";
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            else
            {
                MessageBox.Show("更新程序不可以直接启动！", "PO校验工具", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }



            
        }
    }
}
