using System;
using System.IO;
using Microsoft.Win32;
using System.Drawing;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
//using System.Windows.Data;
//using System.Windows.Documents;
using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Imaging;
//using System.Windows.Navigation;
//using System.Windows.Shapes;

namespace ImageToExcelConverter
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void Button_Convert_Click(object sender, RoutedEventArgs e)
		{
			//Let user open image file, so OpenFileDialog
			OpenFileDialog openFileDlg = new OpenFileDialog();
			openFileDlg.Title = "Open Image file to convert";
			//TODO : Add support to icon files
			openFileDlg.Filter = "Bitmap Files (*.bmp;*.dib)|*.bmp;*.dib" +
								 "|JPEG (*.jpg;*.jpeg;*.jpe;*.jfif)|*.jpg;*.jpeg;*.jpe;*.jfif" +      
								 "|TIFF (*.tif;*.tiff)|*.tif;*.tiff" +
								 "|PNG (*.png)|*.png" +
							   //"|Icon File (*.ico)|*.ico" +
								 "|All Pictures Files|*.bmp;*.dib;*.jpg;*.jpeg;*.jpe;*.jfif;*.tif;*.tiff;*.png;*.ico";
			openFileDlg.FilterIndex = 6;

			//Show the open file dialog
			Nullable<bool> result = openFileDlg.ShowDialog();

			//Config the SaveFileDialog for later use
			SaveFileDialog saveFileDlg = new SaveFileDialog();
			saveFileDlg.FileName = "the excel file";
 			//BIG TODO : Add options about default extension here
			
			
			//BIG TODO : Change to non-interop solution like NPOI if possible, for now interop is ok
			if (result == true)
			{
				//Create a Bitmap with that file, return error if not valid
				try
				{
					Bitmap testBitmap = new Bitmap(openFileDlg.FileName);
				}
				catch (ArgumentException)
				{
					MessageBox.Show("Invalid image file, or format not supported", "Error", 0, MessageBoxImage.Error);
					return;
				}

				Bitmap imageBitmap = new Bitmap(openFileDlg.FileName);

				//Create an Excel instance with interop, create a workbook, and open the 1st worksheet in it
				Excel.Application excelApp = new Excel.Application();
				if (excelApp == null)
				{
					MessageBox.Show("Excel is not properly configured, version " +
						"not compatible, or something happened", "Error", 0, MessageBoxImage.Error);
					return;
				}
				excelApp.Visible = true;
				Excel.Workbook excelBook = excelApp.Workbooks.Add(Missing.Value);
				Excel.Worksheet excelSheet = excelBook.ActiveSheet;

				//Adjust size and make the cells square (represent pixels, square ones at least)
				excelSheet.Cells.ColumnWidth = 0.1;
				excelSheet.Cells.RowHeight = 1;


				//Get the color of each pixel out to the Excel file
				for (int y = 0; y < imageBitmap.Height; y++)
				{
					for (int x = 0; x < imageBitmap.Width; x++)
					{
						//Fix this BUGGGGGGG
						excelSheet.Cells[y+1, x+1].Interior.Color = ColorTranslator.ToOle(imageBitmap.GetPixel(x, y));
					}
				}

				//Show SaveFileDialog
				result = saveFileDlg.ShowDialog();

				//Save document
				if (result == true)
				{
					excelBook.SaveAs(saveFileDlg.FileName);
				}
			}
		}

		private void Button_Cancel_Click(object sender, RoutedEventArgs e)
		{
			System.Windows.Application.Current.Shutdown();
		}
	}
}
