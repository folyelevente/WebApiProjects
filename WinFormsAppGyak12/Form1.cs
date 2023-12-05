using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace WinFormsAppGyak12
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
			CreateExcel();
		}

		Excel.Application xlApp; // A Microsoft Excel alkalmaz�s
		Excel.Workbook xlWB;     // A l�trehozott munkaf�zet
		Excel.Worksheet xlSheet; // Munkalap a munkaf�zeten bel�l

		void CreateExcel()
		{
			try
			{
				// Excel elind�t�sa �s az applik�ci� objektum bet�lt�se
				xlApp = new Excel.Application();

				// �j munkaf�zet
				xlWB = xlApp.Workbooks.Add(Missing.Value);

				// �j munkalap
				xlSheet = xlWB.ActiveSheet;

				// T�bla l�trehoz�sa
				CreateTable(); // Ennek meg�r�sa a k�vetkez� feladatr�szben k�vetkezik

				// Control �tad�sa a felhaszn�l�nak
				xlApp.Visible = true;
				xlApp.UserControl = true;
			}
			catch (Exception ex) // Hibakezel�s a be�p�tett hiba�zenettel
			{
				string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
				MessageBox.Show(errMsg, "Error");

				// Hiba eset�n az Excel applik�ci� bez�r�sa automatikusan
				xlWB.Close(false, Type.Missing, Type.Missing);
				xlApp.Quit();
				xlWB = null;
				xlApp = null;
			}

			void CreateTable()
			{
				string[] fejl�cek = new string[] {
					"K�rd�s",
					"1. v�lasz",
					"2. v�laszl",
					"3. v�lasz",
					"Helyes v�lasz",
					"k�p"
				};

				for (int i = 0; i < fejl�cek.Length; i++)
				{
					xlSheet.Cells[1, 1] = fejl�cek[0];
					Excel.Range fejll�cRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(1, 6);
					fejll�cRange.Font.Bold = true;
					fejll�cRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
					fejll�cRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
					fejll�cRange.EntireColumn.AutoFit();
					fejll�cRange.RowHeight = 40;
					fejll�cRange.Interior.Color = Color.Fuchsia;
					fejll�cRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

					Models.HajosContext hajosContext = new Models.HajosContext();
					var mindenK�rd�s = hajosContext.Questions.ToList();

					object[,] adatT�mb = new object[mindenK�rd�s.Count(), fejl�cek.Count()];

					for (int ii = 0; ii < mindenK�rd�s.Count(); i++)
					{
						adatT�mb[ii, 0] = mindenK�rd�s[i].Question1;
						adatT�mb[ii, 1] = mindenK�rd�s[i].Answer1;
						adatT�mb[ii, 2] = mindenK�rd�s[i].Answer2;
						adatT�mb[ii, 3] = mindenK�rd�s[i].Answer3;
						adatT�mb[ii, 4] = mindenK�rd�s[i].CorrectAnswer;
						adatT�mb[ii, 5] = mindenK�rd�s[i].Image;
					}

					int sorokSz�ma = adatT�mb.GetLength(0);
					int oszlopokSz�ma = adatT�mb.GetLength(1);

					Excel.Range adatRange = xlSheet.get_Range("A2", Type.Missing).get_Resize(sorokSz�ma, oszlopokSz�ma);
					adatRange.Value2 = adatT�mb;

					adatRange.Columns.AutoFit();
				}
			}
		}
	}
}