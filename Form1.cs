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
using ComponentFactory.Krypton.Toolkit;
using Excel = Microsoft.Office.Interop.Excel;
namespace Final_Project
{
    public partial class Form1 : KryptonForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void connect_system(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            int rw = 0;
            int cl = 0;

            string fileName = "user_password.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            var flag1 = false;
            var flag2 = false;
            for(int i = 1; i <= rw; i++)
            {
                flag1 = false;
                flag1 = false;
                for (int j = 1; j <= cl; j++)
                {
                    String str = (string)(range.Cells[i,j] as Excel.Range).Text;
                    if (user_name.Text == str)
                    {
                        flag1 = true;
                    }
                    if (password.Text == str)
                    {
                        flag2 = true;
                    }
                    if (flag1 == true && flag2 == true)
                    {
                        this.SetVisibleCore(false);
                        Form3 frm3 = new Form3();
                        frm3.ShowDialog();
                        return;
                    }
                    else if (flag1 == true && flag2 != true)
                    {
                        string msg = "סיסמה שגוייה";
                        MessageBox.Show(msg);
                        return;
                    }
                   
                }
            }
            string text = "שם משתמש לא קיים";
            MessageBox.Show(text);





        }

        private void sign_up_Click(object sender, EventArgs e)
        {
            
                Form2 frm2 = new Form2();
                frm2.ShowDialog();
            
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
   
    }

