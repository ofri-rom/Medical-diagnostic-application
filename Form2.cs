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
    public partial class Form2 :KryptonForm

    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void sign_up(object sender, EventArgs e)
        {
            int counter_letters_user_name = 0, counter_numbers_user_name = 0, counter_sign_password = 0, counter_letters_password = 0, counter_numbers_password = 0, counter_numbers_id=0;
            for (int i = 0; i < user_name_1.Text.Length; i++)
            {
                if (user_name_1.Text[i] >= 'A' && user_name_1.Text[i] <= 'Z' || user_name_1.Text[i] >= 'a' && user_name_1.Text[i] <= 'z')
                    counter_letters_user_name++;
                if (user_name_1.Text[i] >= '0' && user_name_1.Text[i] <= '9')
                    counter_numbers_user_name++;
            }
            for (int i = 0; i < password.Text.Length; i++)
            {
                if (password.Text[i] >= 'A' && password.Text[i] <= 'Z' || password.Text[i] >= 'a' && password.Text[i] <= 'z')
                    counter_letters_password++;
                if (password.Text[i] >= '0' && password.Text[i] <= '9')
                    counter_numbers_password++;
                if ((password.Text[i] >= '!' && password.Text[i] <= '/') || (password.Text[i] >= ':' && password.Text[i] <= '@') || (password.Text[i] >= '{' && password.Text[i] <= '~'))
                    counter_sign_password++;
            }
            for (int i = 0; i < id.Text.Length; i++)
            {
                if (id.Text[i] >= '0' && id.Text[i] <= '9')
                    counter_numbers_id++;

            }

            if (user_name_1.Text.Length>8 || user_name_1.Text.Length < 6)
            {
                string msg = "שם משתמש חייב להכיל בין 6 ל 8 תווים";
                MessageBox.Show(msg);
                return;
            }
           
            else if (counter_numbers_user_name > 2 || counter_letters_user_name != user_name_1.Text.Length - counter_numbers_user_name)
            {
                string msg = "שם משתמש חייב להכיל לכל היותר 2 ספרות,וכל השאר אותיות";
                MessageBox.Show(msg);
                return;
            }
            else if (password.Text.Length > 10 || password.Text.Length < 8)
            {
                string msg = "סיסמה חייב להכיל בין 8 ל 10 תווים";
                MessageBox.Show(msg);
                return;
            }
            
            else if (counter_numbers_password == 0 || counter_letters_password == 0 || counter_sign_password == 0)
            {
                string msg = "סיסמה חייבת להכיל לפחות אות אחת, סיפרה אחת ותו מיוחד ";
                MessageBox.Show(msg);
                return;
            }
            else if (counter_numbers_id!=9)
            {
                string msg = "ת.ז לא חוקית";
                MessageBox.Show(msg);
                return;
            }

            else
            {
                if (password.Text == validate.Text)
                {
                    Microsoft.Office.Interop.Excel.Application oXL = null;
                    Microsoft.Office.Interop.Excel._Workbook oWB = null;
                    Microsoft.Office.Interop.Excel._Worksheet oSheet = null;
                    Excel.Range range;
                    string fileName = "user_password.xlsx";
                    string path = Path.Combine(Environment.CurrentDirectory, fileName);


                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oWB = oXL.Workbooks.Open(path);
                    oSheet = String.IsNullOrEmpty("s1") ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets["s1"];


                    range = oSheet.UsedRange;
                    int rw = range.Rows.Count;

                    oSheet.Cells[rw + 1, 1] = password.Text;
                    oSheet.Cells[rw + 1, 2] = user_name_1.Text;
                    oSheet.Cells[rw + 1, 3] = id.Text;
                    oWB.Save();
                    oWB.Close();



                    string msg = "נרשמת בהצלחה :) ";
                    MessageBox.Show(msg);
                    this.SetVisibleCore(false);
                }
                else
                {
                    string msg = "סיסמה לא תואמת";
                    MessageBox.Show(msg);
                }

            }

            
            user_name_1.Clear();
            password.Clear();
            validate.Clear();
            





        }

        private void kryptonLabel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void user_name_1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
