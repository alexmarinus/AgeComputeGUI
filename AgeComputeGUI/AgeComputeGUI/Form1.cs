using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
//Merge doar pentru Office 2007!
using Excel = Microsoft.Office.Interop.Excel; 

namespace AgeComputeGUI
{
    public partial class Form1 : Form
    {
        /*Calea completa a fisierului .xlsx pe care se
         * vor efectua calculele*/
        private string fullFilePath;

        /*Determinarea varstei unei persoane in ani, luni si zile
		pornind de la data nasterii si returnarea varstei sub
		forma de sir de caractere*/
        public static string getFullAgeAsString(string birthdate)
        {
            //Data nasterii si data de astazi
            DateTime birthDateObject, todayDateObject;
            //Constructorul de sir de caractere
            StringBuilder stringbuilder = new StringBuilder();
            //Data de astazi
            todayDateObject = DateTime.Now;
            /*Data nasterii*/
            birthDateObject = Convert.ToDateTime(birthdate);
            var timeSpan = todayDateObject - birthDateObject;
            //Diferenta dintre cele doua dati, in zile
            var totaldays = timeSpan.Days;
            //Numarul de ani din diferenta
            var years = totaldays / 365;
            //Numarul de luni din diferenta
            var months = (totaldays - years * 365) / 31;
            //Numarul de zile din diferenta
            var days = totaldays - years * 365 - months * 31;
            //Construirea sirului de caractere pe baza rezultatelor
            stringbuilder.Append(years + " ani " + months + " luni " + days + " zile");
            return stringbuilder.ToString();
        }

        //Initializare fereastra
        public Form1()
        {
            InitializeComponent();
            fullFilePath = "";
        }

        /**Butonul Cautare fisier deschide o fereastra de dialog
         * care permite alegerea fisierului Excel*/
        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                fullFilePath = openFileDialog1.FileName;
            }
        }

        /** Butonul Stergere elibereaza caseta de text */
        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            fullFilePath = "";
        }

        /**Butonul Calcul deschide fisierul dat, parcurge toate
         * intrarile cu datele nasterii si calculeaza varstele
         * pe baza lor. Modificarile sunt aduse in cadrul
         * aceluiasi fisier.*/
        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(fullFilePath))
            {
                MessageBox.Show("Introduceti calea catre fisier!");
                return;
            }

            Excel.Application excelApp = new Excel.Application();
            try
            {
                /*Linie absolut necesara pentru a asigura
                 * accesul la fisier;  acesta este setat ca fiind vizibil
                 de programul apelant*/
                excelApp.Visible = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Eroare la afisare!");
                return;
            }

            /**Se deschide fisierul .xlsx (workbook) specificat*/
            Excel.Workbook excelWorkbook;
            Excel.Worksheet excelWorksheet;
            Excel.Range excelRange;
            object misValue = System.Reflection.Missing.Value;

            try
            {
                excelWorkbook = excelApp.Workbooks.Open(@fullFilePath);
            }
            catch (Exception)
            {
                MessageBox.Show("Nu s-a putut deschide fisierul specificat!");
                return;
            }
            /*Se specifica faptul ca sheet-ul de ineteres este primul*/
            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);

            /*Se determina numarul ultimului rand populat din worksheet*/
            excelRange = (Excel.Range)excelWorksheet.Cells[excelWorksheet.Rows.Count, 1];
            int lastFullRow = (int) excelRange.get_End(Excel.XlDirection.xlUp).Row;

            /*Numarul unui rand din worksheet*/
            int row;
            string birthdate;

            /*Lista cu varstele determinate pentru toate persoanele*/
            List<string> ages = new List<string>();

            /*Se parcurg casutele din tabela ce contin datile de nastere.
             *Varstele se adauga in lista de mai sus.
             *Parcurgerea se incepe de la randul 2, deoarece randul 1
             *reprezinta capul tabelului (nu contine date efective)*/
            for (row = 2; row <= lastFullRow; row ++) {
                birthdate = excelWorksheet.Cells[row, 2].Value.ToString();
                ages.Add (getFullAgeAsString(birthdate));
            }

            /*Inchid workbook-ul (fisierul Excel)*/
            excelWorkbook.Close(true, misValue, misValue);
            excelApp.Quit();

            /*Il redeschid pentru a scrie rezultatele obtinute
             * in coloana Varsta*/
            excelWorkbook = excelApp.Workbooks.Open(@fullFilePath);
            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);

            int index;
            /*Lista fiind numerotata de la 0 si randurile worksheet-ului
             * numerotate la 1, sirul cu indicele index din lista ages
             * se va regasi pe randul index + 2 din worksheet*/
            for (index = 0; index < ages.Count; index++)
                excelWorksheet.Cells[index + 2, 3] = ages[index];

            excelWorkbook.Close(true, misValue, misValue);
            excelApp.Quit();

            /*Cand s-a terminat scrierea rezultatelor, se
             * goleste lista de varste*/
            ages.Clear();

            /*Se sterg referintele la obiectele ce tin de
             * fisierul Excel*/
            Marshal.ReleaseComObject(excelWorksheet);
            Marshal.ReleaseComObject(excelWorkbook);
            Marshal.ReleaseComObject(excelApp);
        }

        /**Butonul Iesire inchide aplicatia */
        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
