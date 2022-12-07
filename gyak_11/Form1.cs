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

            Excel.Application xlApp; // A Microsoft Excel alkalmazás
            Excel.Workbook xlWB;     // A létrehozott munkafüzet
            Excel.Worksheet xlSheet; // Munkalap a munkafüzeten belül

            try
            {
                // Excel elindítása és az applikáció objektum betöltése
                xlApp = new Excel.Application();

                // Új munkafüzet
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                // Új munkalap
                xlSheet = xlWB.ActiveSheet;

                // Tábla létrehozása
                CreateTable(xlSheet); // Ennek megírása a következõ feladatrészben következik

                // Control átadása a felhasználónak
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) // Hibakezelés a beépített hibaüzenettel
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                // Hiba esetén az Excel applikáció bezárása automatikusan
                xlWB = null;
                xlApp = null;
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
            }
        }

        void CreateTable(Excel.Worksheet shit)
        {
            string[] fejlécek = new string[] {
                "Kérdés",
                "1. válasz",
                "2. válaszl",
                "3. válasz",
                "Helyes válasz",
                "kép"};
            for (int i = 0; i < fejlécek.Length; i++)
            {
                shit.Cells[1, i+1] = fejlécek[i];
            }

            var mindenKérdés = context.Questions.ToList();
            object[,] adat = new object[mindenKérdés.Count(), fejlécek.Count()];

            for (int i = 0; i < mindenKérdés.Count(); i++)
            {
                adat[i, 0] = mindenKérdés[i].Question1;
                adat[i, 1] = mindenKérdés[i].Answer1;
                adat[i, 2] = mindenKérdés[i].Answer2;
                adat[i, 3] = mindenKérdés[i].Answer3;
                adat[i, 4] = mindenKérdés[i].CorrectAnswer;
                adat[i, 5] = mindenKérdés[i].Image;
            }

            int sorokSzáma = adat.GetLength(0);
            int oszlopokSzáma = adat.GetLength(1);

            Excel.Range adatRange = shit.get_Range("A2", Type.Missing).get_Resize(sorokSzáma, oszlopokSzáma);
            adatRange.Value2 = adat;

            adatRange.Columns.AutoFit();

            Excel.Range fejllécRange = shit.get_Range("A1", Type.Missing).get_Resize(1, 6);
            fejllécRange.Font.Bold = true;
            fejllécRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            fejllécRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            fejllécRange.EntireColumn.AutoFit();
            fejllécRange.RowHeight = 40;
            fejllécRange.Interior.Color = Color.Fuchsia;
            fejllécRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            int lastRowID = shit.UsedRange.Rows.Count;

            Excel.Range teljes = shit.get_Range("A1", Type.Missing).get_Resize(lastRowID,oszlopokSzáma);
            teljes.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range elso_o = shit.get_Range("A2", Type.Missing).get_Resize(lastRowID, 1);
            elso_o.Font.Bold = true;
            elso_o.Interior.Color = Color.LightYellow;

            Excel.Range utolso_o = shit.get_Range("F2", Type.Missing).get_Resize(lastRowID, 1);
            utolso_o.Interior.Color = Color.LightGreen;
        }
    }
}