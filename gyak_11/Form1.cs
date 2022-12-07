using gyak_11.Models;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace gyak_11
{
    public partial class Form1 : Form
    {
        HajosContext context = new HajosContext();
        public Form1()
        {
            InitializeComponent();

            Excel.Application xlApp; // A Microsoft Excel alkalmaz�s
            Excel.Workbook xlWB;     // A l�trehozott munkaf�zet
            Excel.Worksheet xlSheet; // Munkalap a munkaf�zeten bel�l

            try
            {
                // Excel elind�t�sa �s az applik�ci� objektum bet�lt�se
                xlApp = new Excel.Application();

                // �j munkaf�zet
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                // �j munkalap
                xlSheet = xlWB.ActiveSheet;

                // T�bla l�trehoz�sa
                CreateTable(xlSheet); // Ennek meg�r�sa a k�vetkez� feladatr�szben k�vetkezik

                // Control �tad�sa a felhaszn�l�nak
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) // Hibakezel�s a be�p�tett hiba�zenettel
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                // Hiba eset�n az Excel applik�ci� bez�r�sa automatikusan
                xlWB = null;
                xlApp = null;
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
            }
        }

        void CreateTable(Excel.Worksheet shit)
        {
            string[] fejl�cek = new string[] {
                "K�rd�s",
                "1. v�lasz",
                "2. v�laszl",
                "3. v�lasz",
                "Helyes v�lasz",
                "k�p"};
            for (int i = 0; i < fejl�cek.Length; i++)
            {
                shit.Cells[1, i+1] = fejl�cek[i];
            }

            var mindenK�rd�s = context.Questions.ToList();
            object[,] adat = new object[mindenK�rd�s.Count(), fejl�cek.Count()];

            for (int i = 0; i < mindenK�rd�s.Count(); i++)
            {
                adat[i, 0] = mindenK�rd�s[i].Question1;
                adat[i, 1] = mindenK�rd�s[i].Answer1;
                adat[i, 2] = mindenK�rd�s[i].Answer2;
                adat[i, 3] = mindenK�rd�s[i].Answer3;
                adat[i, 4] = mindenK�rd�s[i].CorrectAnswer;
                adat[i, 5] = mindenK�rd�s[i].Image;
            }

            int sorokSz�ma = adat.GetLength(0);
            int oszlopokSz�ma = adat.GetLength(1);

            Excel.Range adatRange = shit.get_Range("A2", Type.Missing).get_Resize(sorokSz�ma, oszlopokSz�ma);
            adatRange.Value2 = adat;

            adatRange.Columns.AutoFit();

            Excel.Range fejll�cRange = shit.get_Range("A1", Type.Missing).get_Resize(1, 6);
            fejll�cRange.Font.Bold = true;
            fejll�cRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            fejll�cRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            fejll�cRange.EntireColumn.AutoFit();
            fejll�cRange.RowHeight = 40;
            fejll�cRange.Interior.Color = Color.Fuchsia;
            fejll�cRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            int lastRowID = shit.UsedRange.Rows.Count;

            Excel.Range teljes = shit.get_Range("A1", Type.Missing).get_Resize(lastRowID,oszlopokSz�ma);
            teljes.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range elso_o = shit.get_Range("A2", Type.Missing).get_Resize(lastRowID, 1);
            elso_o.Font.Bold = true;
            elso_o.Interior.Color = Color.LightYellow;

            Excel.Range utolso_o = shit.get_Range("F2", Type.Missing).get_Resize(lastRowID, 1);
            utolso_o.Interior.Color = Color.LightGreen;
        }
    }
}