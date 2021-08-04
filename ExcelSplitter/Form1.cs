using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ExcelSplitter
{
	public partial class Form1 : Form
	{
		private string startPath = @"C:\";
		private string filePath_source = @"C:\";
		private string filePath_copiedinto = Path.GetTempPath() + @"temp.xlsx";
		private string destinationPath = @"C:\";

		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			log("Splitter initialized");
		}

		// start splitting excel button
		private void button1_Click(object sender, EventArgs e)
		{
			log("Splitting process starting");
			log("Creating temp excel file to " + filePath_copiedinto);
			var app = new Microsoft.Office.Interop.Excel.Application();
			app.DisplayAlerts = false;
			var wb = app.Workbooks.Add(Type.Missing);
			wb.SaveCopyAs(filePath_copiedinto);
			wb.Close();
			backgroundWorker1.RunWorkerAsync();
		}



		// select excel file dialog box button
		private void button2_Click(object sender, EventArgs e)
		{
			OpenFileDialog excelFileBrowserDialog = new OpenFileDialog();
			excelFileBrowserDialog.InitialDirectory = @"C:\";
			excelFileBrowserDialog.Filter = "Excel files|*.xlsx";
			excelFileBrowserDialog.ShowDialog();
			filePath_source = excelFileBrowserDialog.FileName;
			textBox1.Text = filePath_source;
			log("Excel file selected");
		}

		// select destination folder dialog box button
		private void button3_Click(object sender, EventArgs e)
		{
			FolderBrowserDialog destinationPathBrowserDialog = new FolderBrowserDialog();
			destinationPathBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop;
			destinationPathBrowserDialog.ShowDialog();
			destinationPath = destinationPathBrowserDialog.SelectedPath;
			textBox3.Text = destinationPath;
			log("Destination path selected");
		}

		// excel file splitting background worker
		private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
		{
			try
			{
				Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
				app.DisplayAlerts = false;
				Microsoft.Office.Interop.Excel.Workbook book = app.Workbooks.Open(filePath_source);
				Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Worksheets.get_Item(1);

				int iRowCount = sheet.UsedRange.Rows.Count;
				int maxrows = int.Parse(textBox2.Text);
				int maxloops = iRowCount / maxrows;
				int beginrow = 2; //skipping the header row.

				Microsoft.Office.Interop.Excel.Application destxlApp;
				Microsoft.Office.Interop.Excel.Workbook destworkBook;
				Microsoft.Office.Interop.Excel.Worksheet destworkSheet;
				Microsoft.Office.Interop.Excel.Range destrange;



				for (int i = 1; i <= maxloops; i++)
				{
					log("Starting to work for file number: " + i + " and starting row # from: " + beginrow);
					Microsoft.Office.Interop.Excel.Range startCell = sheet.Cells[beginrow, 1];
					Microsoft.Office.Interop.Excel.Range
						endCell = sheet.Cells[beginrow + maxrows - 1, 3]; 
					Microsoft.Office.Interop.Excel.Range rng = sheet.Range[startCell, endCell];
					rng = rng.EntireRow; 
					rng.Copy(Type.Missing);

					//opening of the second worksheet and pasting
					string destPath = filePath_copiedinto;
					destxlApp = new Microsoft.Office.Interop.Excel.Application();
					destxlApp.DisplayAlerts = true;
					destxlApp.Visible = true;
					destworkBook = destxlApp.Workbooks.Open(destPath, 0, false);
					destworkSheet = destworkBook.Worksheets.get_Item(1);
					destrange = destworkSheet.Cells[2, 1];
					destrange.Select();
					Thread.Sleep(1000);
					destworkSheet.Paste(Type.Missing, Type.Missing);
					Thread.Sleep(1000);
					destworkBook.SaveAs(destinationPath + "\\" + i + ".xlsx");
					Thread.Sleep(1000);
					beginrow = beginrow + maxrows;
					backgroundWorker1.ReportProgress(i, maxloops);

					GC.Collect();
					GC.WaitForPendingFinalizers();

					Marshal.FinalReleaseComObject(destrange);
					Marshal.FinalReleaseComObject(destworkSheet);

					destworkBook.Close(true, null, null);
					Marshal.FinalReleaseComObject(destworkBook);

					destxlApp.Quit();
					Marshal.FinalReleaseComObject(destxlApp);
				}
			}
			catch (Exception ex)
			{
				backgroundWorker1.CancelAsync();
				log("Splitting process failed: " + ex.Message);
			}
		}

		private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
		{
			int maxLoops = (int) e.UserState;
			progressBar1.Value = (e.ProgressPercentage / maxLoops)*100;
			label8.Text = "Overall Progress %" + (e.ProgressPercentage / maxLoops) * 100;
			label5.Text = e.ProgressPercentage.ToString();
		}

		private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
		{
			log("Splitting process completed");
		}

		private void log(string log)
		{
			string formattedLogText = DateTime.Now.ToShortTimeString() + " ---- " + log;
			BeginInvoke((MethodInvoker)delegate
			{
				richTextBox1.Text += formattedLogText + Environment.NewLine;
			});
		}
	}
}
