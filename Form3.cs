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
    public partial class Form3 : KryptonForm
    {
        static int rows=2;
        string[] fields = new string[22];
        public Form3()
        {
            
            InitializeComponent();
            kryptonPanel2.Visible = true;
            inser_value_panel.Visible = false;
            diagnostic_panel.Visible = false;
            
        }

        private void insert_values_Click(object sender, EventArgs e)
        {
            kryptonPanel2.Visible = false;
            diagnostic_panel.Visible = false;
            inser_value_panel.Visible = true;
        }

        private void kryptonTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void kryptonLabel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void exit_button(object sender, EventArgs e)
        {
            this.Close();
        }

        private void init_array(object sender, EventArgs e)
        {
            if (first_name.Text.All(char.IsLetter)==false)
            {
                string first_name_error = "שם פרטי לא יכול להכיל מספרים";
                MessageBox.Show(first_name_error);
                return;
            }
            fields[0] = first_name.Text;
            if (Last_name.Text.All(char.IsLetter) == false)
            {
                string last_name_error = "שם משפחה לא יכול להכיל מספרים";
                MessageBox.Show(last_name_error);
                return;
            }
            fields[1] = Last_name.Text;
            if(ID.Text.Length>9 || ID.Text.Length < 9)
            {
                string id_error = "תעודת זהות לא חוקית";
                MessageBox.Show(id_error);
                return;
            }
            fields[2] = ID.Text;
            if (Int32.Parse(age.Text) < 0 || Int32.Parse(age.Text) > 100)
            {
                string age_error = "גיל אינו חוקי";
                MessageBox.Show(age_error);
                return;
            }
            fields[3] = age.Text;
            fields[4] = WBC.Text;
            fields[5] = Neut.Text;
            fields[6] = Lymph.Text;
            fields[7] = RBC.Text;
            fields[8] = HCT.Text;
            fields[9] = Urea.Text;
            fields[10] = Hb.Text;
            fields[11] = Crtn.Text;
            fields[12] = Iron.Text;
            fields[13] = HDL.Text;
            fields[14] = AP.Text;
            fields[15] = Smoke.Text;
            fields[16] = east_choose.Text;
            fields[17] = prgnant_choose.Text;
            fields[18] = diarea.Text;
            fields[19] = temperature.Text;
            fields[20] = gender.Text;
            fields[21] = ethopia.Text;
            string msg = "הפרטים נקלטו בהצלחה";
            Write_to_excel();
            MessageBox.Show(msg);
        }

        private void inser_value_panel_Paint(object sender, PaintEventArgs e)
        {

        }
        public string recomendation(string disease)
        {
            if (disease == "אנמיה")
                return "ביום למשך חודש" + "B12" + "שני כדורי 10 מג של";
            else if (disease == "דיאטה")
                return "לתאם פגישה עם תזונאי";
            else if (disease == "דימום")
                return "להתפנות בדחיפות לבית החולים";
            else if (disease == "היפרליפידמיה")
                return "לתאם פגישה עם תזונאי, כדור 5 מג של סימוביל ביום למשך שבוע";
            else if (disease == "הפרעה ביצירת הדם / תאי דם")
                return "כדור 5 מג של חומצה פולית ביום למשך חודש" + "\n" + "ביום למשך חודש" + "B12" + "כדור 10 מג של";
            else if (disease == "הפרעה המטולוגית")
                return "זריקה של הורמון לעידוד ייצור תאי הדם האדומים";
            else if (disease == "הרעלת ברזל")
                return "להתפנות לבית החולים";
            else if (disease == "התייבשות")
                return "מנוחה מוחלטת בשכיבה, החזרת נוזלים בשתייה";
            else if (disease == "זיהום")
                return "אנטיביוטיקה ייעודית";
            else if (disease == "חוסר בוויטמינים")
                return "הפנייה לבדיקת דם לזיהוי הוויטמינים החסרים";
            else if (disease == "מחלה ויראלית")
                return "לנוח בבית";
            else if (disease == "מחלות בדרכי המרה")
                return "הפנייה לטיפול כירורגי";
            else if (disease == "מחלות לב")
                return "לתאם פגישה עם תזונאי";
            else if (disease == "מחלת דם")
                return "שילוב של ציקלופוספאמיד וקורטיקוסרואידים";
            else if (disease == "מחלת כבד")
                return "הפנייה לאבחנה ספציפית לצורך קביעת טיפול";

            else if (disease == "מחלת כליה")
                return "איזון את רמות הסוכר בדם";
            else if (disease == "מחסור בברזל")
                return "ביום למשך חודש" + "B12" + "שני כדורי 10 מג של";
            else if (disease == "מחלות שריר")
                return "של אלטמן ביום למשך חודש" + "c3" + "שני כדורי 5 מג של כורכום";
            else if (disease == "מעשנים")
                return "להפסיק לעשן";
            else if (disease == "מחלת ריאות")
                return "להפסיק לעשן / הפנייה לצילום רנטגן של הריאות";
            else if (disease == "פעילות יתר של בלוטת התריס")
                return "ropylthiouracil להקטנת פעילות בלוטת התריס ";
            else if (disease == "סוכרת מבוגרים")
                return "התאמת אינסולין למטופל";
            else if (disease == "סרטן")
                return "Entrectinib-אנטרקטיניב";
            else if (disease == "צריכה מוגברת של בשר")
                return "לתאם פגישה עם תזונאי";
            else if (disease == "שימוש בתרופות שונות")
                return "הפנייה לרופא המשפחה לצורך בדיקת התאמה בין התרופות";
            else if (disease == "תת תזונה")
                return "לתאם פגישה עם תזונאי";
            return "No recomandation";
        }

        public string[] disease()
        {
            int cuonter = 0;
            string[] d = new string[26];
            if (((float.Parse(fields[3])) <= 3 && (float.Parse(fields[3])) >= 0 && (float.Parse(fields[4])) > 17500) || ((float.Parse(fields[3])) <= 17 && (float.Parse(fields[3])) >= 4 && (float.Parse(fields[4])) > 15500) || ((float.Parse(fields[3])) >= 18 && (float.Parse(fields[4])) > 11000))
                if (fields[19] == "yes")
                {
                    if (!d.Contains("זיהום"))
                    {
                        d[cuonter] = "זיהום";
                        cuonter++;
                    }
                }

                else
                {
                    if (!d.Contains("מחלת דם"))
                    {
                        d[cuonter] = "מחלת דם";
                        cuonter++;
                    }
                    if (!d.Contains("סרטן"))
                    {
                        d[cuonter] = "סרטן";
                        cuonter++;
                    }

                }
            else if (((float.Parse(fields[3])) <= 3 && (float.Parse(fields[3])) >= 0 && (float.Parse(fields[4])) < 6000) || ((float.Parse(fields[3])) <= 17 && (float.Parse(fields[3])) >= 4 && (float.Parse(fields[4])) < 5500) || ((float.Parse(fields[3])) >= 18 && (float.Parse(fields[4])) < 4500))
            {
                if (!d.Contains("ויראלית"))
                {
                    d[cuonter] = "ויראלית";
                    cuonter++;
                }
                if (!d.Contains("סרטן"))
                {
                    d[cuonter] = "סרטן";
                    cuonter++;
                }

            }
            if ((100 * (float.Parse(fields[5])) / (float.Parse(fields[4]))) > 54)
            {
                if (!d.Contains("זיהום"))
                {
                    d[cuonter] = "זיהום";
                    cuonter++;
                }

            }
            else if ((100 * (float.Parse(fields[5])) / (float.Parse(fields[4]))) < 28)
            {
                if (!d.Contains("הפרעה ביצירת הדם / תאי דם"))
                {

                    d[cuonter] = "הפרעה ביצירת הדם / תאי דם";
                    cuonter++;
                }
                if (!d.Contains("זיהום"))
                {
                    d[cuonter] = "זיהום";
                    cuonter++;
                }
                if (!d.Contains("סרטן"))
                {
                    d[cuonter] = "סרטן";
                    cuonter++;
                }

            }
            if ((100 * (float.Parse(fields[6])) / (float.Parse(fields[4]))) > 52)
            {
                if (!d.Contains("זיהום"))
                {
                    d[cuonter] = "זיהום";
                    cuonter++;
                }
                if (!d.Contains("סרטן"))
                {
                    d[cuonter] = "סרטן";
                    cuonter++;
                }
            }
            else if ((100 * (float.Parse(fields[6])) / (float.Parse(fields[4]))) < 36)
            {
                if (!d.Contains("הפרעה ביצירת הדם / תאי דם"))
                {
                    d[cuonter] = "הפרעה ביצירת הדם / תאי דם";
                    cuonter++;
                }
            }
            if ((float.Parse(fields[7])) > 6)
            {
                if (!d.Contains("הפרעה ביצירת הדם / תאי דם"))
                {
                    d[cuonter] = "הפרעה ביצירת הדם / תאי דם";
                    cuonter++;
                }
                if (fields[15] == "yes")
                {
                    if (!d.Contains("מעשנים"))
                    {
                        d[cuonter] = "מעשנים";
                        cuonter++;
                    }

                }
            }
            else if ((float.Parse(fields[7])) < 4.5)
            {
                if (!d.Contains("אנמיה"))
                {
                    d[cuonter] = "אנמיה";
                    cuonter++;
                }
                if (!d.Contains("דימום"))
                {
                    d[cuonter] = "דימום";
                    cuonter++;
                }

            }
            if ((( 100* (float.Parse(fields[8]))) / ((float.Parse(fields[4])) + (float.Parse(fields[7]))) > 54 && fields[20] == "גבר") || (( 100* (float.Parse(fields[8]))) / ((float.Parse(fields[4])) + (float.Parse(fields[7]))) > 47 && fields[20] == "אישה"))
            {

                if (!d.Contains("מעשנים"))
                {
                    d[cuonter] = "מעשנים";
                    cuonter++;
                }
            }
            else if (((((float.Parse(fields[4])) + (float.Parse(fields[7]))) * (float.Parse(fields[8]))) / 100 < 37 && fields[20] == "גבר") || ((((float.Parse(fields[4])) + (float.Parse(fields[7]))) * (float.Parse(fields[8]))) / 100 < 33 && fields[20] == "אישה"))
            {
                if (!d.Contains("אנמיה"))
                {
                    d[cuonter] = "אנמיה";
                    cuonter++;
                }
                if (!d.Contains("דימום"))
                {
                    d[cuonter] = "דימום";
                    cuonter++;
                }
            }
            if ((float.Parse(fields[9])) < 17 && fields[16] == "no")
            {
                if (!d.Contains("תת תזונה"))
                {
                    d[cuonter] = "תת תזונה";
                    cuonter++;
                }
                if (!d.Contains("מחלת כבד"))
                {
                    d[cuonter] = "מחלת כבד";
                    cuonter++;
                }
                if (!d.Contains("דיאטה"))
                {
                    d[cuonter] = "דיאטה";
                    cuonter++;
                }

            }
            else if ((float.Parse(fields[9])) < 18.7 && fields[16] == "yes")
            {
                if (!d.Contains("תת תזונה"))
                {
                    d[cuonter] = "תת תזונה";
                    cuonter++;
                }
                if (!d.Contains("מחלת כבד"))
                {
                    d[cuonter] = "מחלת כבד";
                    cuonter++;
                }
                if (!d.Contains("דיאטה"))
                {
                    d[cuonter] = "דיאטה";
                    cuonter++;
                }

            }
            else if ((float.Parse(fields[9])) > 43 && fields[16] == "no")
            {
                if (!d.Contains("התייבשות"))
                {
                    d[cuonter] = "התייבשות";
                    cuonter++;
                }
                if (!d.Contains("מחלת כליה"))
                {
                    d[cuonter] = "מחלת כליה";
                    cuonter++;
                }
                if (!d.Contains("דיאטה"))
                {
                    d[cuonter] = "דיאטה";
                    cuonter++;
                }

            }
            else if ((float.Parse(fields[9])) > 47.3 && fields[16] == "yes")
            {

                if (!d.Contains("התייבשות"))
                {
                    d[cuonter] = "התייבשות";
                    cuonter++;
                }
                if (!d.Contains("מחלת כליה"))
                {
                    d[cuonter] = "מחלת כליה";
                    cuonter++;
                }
                if (!d.Contains("דיאטה"))
                {
                    d[cuonter] = "דיאטה";
                    cuonter++;
                }
            }
            if (((float.Parse(fields[10])) < 12 && (float.Parse(fields[3])) >= 18) || ((float.Parse(fields[3])) <= 17 && (float.Parse(fields[3])) >= 0 && (float.Parse(fields[10])) < 11.5))
            {
                if (!d.Contains("אנמיה"))
                {
                    d[cuonter] = "אנמיה";
                    cuonter++;
                }
                if (!d.Contains("מחסור בברזל"))
                {
                    d[cuonter] = "מחסור בברזל";
                    cuonter++;
                }
                if (!d.Contains("דימום"))
                {
                    d[cuonter] = "דימום";
                    cuonter++;
                }

            }
            if (((float.Parse(fields[3])) <= 2 && (float.Parse(fields[3])) >= 0 && (float.Parse(fields[11])) > 0.5) || ((float.Parse(fields[3])) <= 17 && (float.Parse(fields[3])) >= 3 && (float.Parse(fields[11])) > 1) || ((float.Parse(fields[3])) >= 18 && (float.Parse(fields[3])) <= 59 && (float.Parse(fields[11])) > 1) || ((float.Parse(fields[3])) >= 60 && (float.Parse(fields[11])) > 1.2))
            {
                if (!d.Contains("מחלת כליה"))
                {
                    d[cuonter] = "מחלת כליה";
                    cuonter++;
                }
                if (!d.Contains("צריכה מוגברת של בשר"))
                {
                    d[cuonter] = "צריכה מוגברת של בשר";
                    cuonter++;
                }

            }
            else if (((float.Parse(fields[3])) <= 2 && (float.Parse(fields[3])) >= 0 && (float.Parse(fields[11])) < 0.2) || ((float.Parse(fields[3])) <= 17 && (float.Parse(fields[3])) >= 3 && (float.Parse(fields[11])) < 0.5) || ((float.Parse(fields[3])) >= 18 && (float.Parse(fields[3])) <= 59 && (float.Parse(fields[11])) < 0.6) || ((float.Parse(fields[3])) >= 60 && (float.Parse(fields[11])) < 0.6))
            {
                if (!d.Contains("תת תזונה"))
                {
                    d[cuonter] = "תת תזונה";
                    cuonter++;
                }


            }
            if (((float.Parse(fields[12])) > 160 && fields[20] == "גבר") || ((float.Parse(fields[12])) > 128 && fields[20] == "אישה"))
            {
                if (!d.Contains("הרעלת ברזל"))
                {
                    d[cuonter] = "הרעלת ברזל";
                    cuonter++;
                }

            }
            else if (((float.Parse(fields[12])) < 60 && fields[20] == "גבר") || ((float.Parse(fields[12])) < 48 && fields[20] == "אישה"))
            {
                if (!d.Contains("תת תזונה"))
                {
                    d[cuonter] = "תת תזונה";
                    cuonter++;
                }
                if (!d.Contains("מחסור בברזל"))
                {
                    d[cuonter] = "מחסור בברזל";
                    cuonter++;
                }
                if (!d.Contains("דימום"))
                {
                    d[cuonter] = "דימום";
                    cuonter++;
                }


            }
            if (((float.Parse(fields[13])) < 34.8 && fields[20] == "גבר" && fields[21] == "yes") || ((float.Parse(fields[13])) < 40.8 && fields[20] == "אישה" && fields[21] == "yes") || ((float.Parse(fields[13])) < 29 && fields[20] == "גבר" && fields[21] == "no") || ((float.Parse(fields[13])) < 34 && fields[20] == "אישה" && fields[21] == "no"))
            {
                if (!d.Contains("סוכרת מבוגרים"))
                {
                    d[cuonter] = "סוכרת מבוגרים";
                    cuonter++;
                }
                if (!d.Contains("מחלות לב"))
                {
                    d[cuonter] = "מחלות לב";
                    cuonter++;
                }
                if (!d.Contains("היפרליפידמיה"))
                {
                    d[cuonter] = "היפרליפידמיה";
                    cuonter++;
                }

            }
            if (((float.Parse(fields[14])) > 120 && fields[16] == "yes") || ((float.Parse(fields[14])) > 90 && fields[16] == "no"))
            {
                if (!d.Contains("מחלות בדרכי המרה"))
                {
                    d[cuonter] = "מחלות בדרכי המרה";
                    cuonter++;
                }
                if (!d.Contains("מחלת כבד"))
                {
                    d[cuonter] = "מחלת כבד";
                    cuonter++;
                }
                if (!d.Contains("פעילות יתר של בלוטת התריס"))
                {
                    d[cuonter] = "פעילות יתר של בלוטת התריס";
                    cuonter++;
                }
                if (!d.Contains("שימוש בתרופות שונות"))
                {
                    d[cuonter] = "שימוש בתרופות שונות";
                    cuonter++;
                }

            }
            else if (((float.Parse(fields[14])) < 60 && fields[16] == "yes") || ((float.Parse(fields[14])) < 30 && fields[16] == "no"))
            {
                if (!d.Contains("חוסר בוויטמינים"))
                {
                    d[cuonter] = "חוסר בוויטמינים";
                    cuonter++;
                }
                if (!d.Contains("תת תזונה"))
                {
                    d[cuonter] = "תת תזונה";
                    cuonter++;
                }

            }
            return d;
        }

        private void diagnostic_test(object sender, EventArgs e)
        {
            if (fields[0] == null)
            {
                string msg = "לא ניתן לאבחן מטופל ללא הכנסת פרטים";
                MessageBox.Show(msg);
                return;
            }
            listView1.Items.Clear();
            inser_value_panel.Visible = false;
            diagnostic_panel.Visible = true;
            string[] test = disease();
            string[] test2 = new string[test.Length];
            for (int i = 0; i < test.Length; i++)
            {
                if (test[i] != null) {
                    test2[i] = recomendation(test[i]);
                    ListViewItem item = new ListViewItem();
                    item.SubItems[0].Text = test2[i];
                    item.SubItems.Add(test[i]);
                    listView1.Items.Add(item);
                }   
            }
        }

        private void import_file(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Filter = "Excel files (*.xlsx)|*.xlsx|Excel files(*.xls)|*.xls";
            openfile.FilterIndex = 2;
            openfile.RestoreDirectory = true;
            openfile.ShowDialog();
            string path = openfile.FileName;



            Microsoft.Office.Interop.Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;


            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            cl = range.Columns.Count;
            for (int i = 1; i <= cl; i++)
            {
                String str = (string)(range.Cells[2, i] as Excel.Range).Text;
                fields[i + 3] = str;
            }
            WBC.Text = fields[4];
            Neut.Text = fields[5];
            Lymph.Text = fields[6];
            RBC.Text = fields[7];
            HCT.Text = fields[8];
            Urea.Text = fields[9];
            Hb.Text = fields[10];
            Crtn.Text = fields[11];
            Iron.Text = fields[12];
            HDL.Text = fields[13];
            AP.Text = fields[14];
        }
        public void Write_to_excel()
        {
            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            Microsoft.Office.Interop.Excel._Worksheet oSheet = null;
            Excel.Range range;

            string fileName = "output.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);

            oXL = new Microsoft.Office.Interop.Excel.Application();
            oWB = oXL.Workbooks.Open(path);
            oSheet = String.IsNullOrEmpty("s1") ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets["s1"];


            range = oSheet.UsedRange;
            int cl = range.Columns.Count;

            for(int i = 1; i <= cl-2; i++)
                oSheet.Cells[rows, i] = fields[i-1];
            if (fields[0] == "")
            {
                string msg = "לא ניתן לאבחן מטופל ללא הכנסת פרטים";
                MessageBox.Show(msg);
                return;
            }
            string[] test = disease();
            string[] test2 = new string[test.Length];
            for (int i = 0; i < test.Length; i++)
            {
                if (test[i] != null)
                {
                    test2[i] = recomendation(test[i]);
                    oSheet.Cells[rows, cl-1] = test[i];
                    oSheet.Cells[rows, cl] = test2[i];
                    rows++;

                }
            }

            oWB.Save();
            oWB.Close();
        }

        private void kryptonLabel26_Paint(object sender, PaintEventArgs e)
        {

        }

        private void richTextBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void diagnostic_panel_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
