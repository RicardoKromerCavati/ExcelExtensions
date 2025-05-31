using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Schema;

namespace ExcelExtensions.Forms;

public partial class OpenFileDialogForm : Form
{
	private OpenFileDialog _openFileDialog;

	public OpenFileDialogForm()
	{
		InitializeComponent();
	}

	private void button1_Click(object sender, EventArgs e)
	{
		_openFileDialog = new OpenFileDialog
		{
			FileName = "Selecione um arquivo Excel",
			Filter = "Arquivos Excel (*.xlsx)|*.xlsx" 
		};


		if (_openFileDialog.ShowDialog() == DialogResult.OK)
		{
			try
			{
				var fileName = _openFileDialog.FileName;

				using (var fStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
				{
					using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fStream, false))
					{
						WorkbookPart workbookPart = doc.WorkbookPart;
						SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
						SharedStringTable sst = sstpart.SharedStringTable;

						WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
						Worksheet sheet = worksheetPart.Worksheet;

						var cells = sheet.Descendants<Cell>();
						var rows = sheet.Descendants<Row>();

						Console.WriteLine("Row count = {0}", rows.LongCount());
						Console.WriteLine("Cell count = {0}", cells.LongCount());

						// One way: go through each cell in the sheet
						//foreach (Cell cell in cells)
						//{
						//	if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
						//	{
						//		int ssid = int.Parse(cell.CellValue.Text);
						//		string str = sst.ChildElements[ssid].InnerText;
						//		Console.WriteLine("Shared string {0}: {1}", ssid, str);
						//	}
						//	else if (cell.CellValue != null)
						//	{
						//		Console.WriteLine("Cell contents: {0}", cell.CellValue.Text);
						//	}
						//}
						
						var directory = Path.GetDirectoryName(fileName);

						var csvFilePath = Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(fileName)}.csv");

						using var file = File.OpenWrite(csvFilePath);

						using var streamWriter = new StreamWriter(file);

						// Or... via each row
						foreach (Row row in rows)
						{
							var line = string.Empty;

							var cellElements = row.Elements<Cell>().ToArray();

							var length = cellElements.Length;

							for (var i = 0; i < length; i++)
							{
								var c = cellElements[i];
								if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
								{
									int ssid = int.Parse(c.CellValue.Text);
									string str = sst.ChildElements[ssid].InnerText;
									Console.WriteLine("Shared string {0}: {1}", ssid, str);

									if (i == length - 1)
									{
										line += $"{str}";
									}
									else
									{
										line += $"{str},";
									}

								}
								else if (c.CellValue != null)
								{
									Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
								}
							}

							//foreach (Cell c in cellElements)
							//{
							//	if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
							//	{
							//		int ssid = int.Parse(c.CellValue.Text);
							//		string str = sst.ChildElements[ssid].InnerText;
							//		Console.WriteLine("Shared string {0}: {1}", ssid, str);
							//		line += $"{str},";
							//	}
							//	else if (c.CellValue != null)
							//	{
							//		Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
							//	}
							//}

							streamWriter.WriteLine(line);
							streamWriter.WriteLine();
						}
					}
				}

				MessageBox.Show($"Arquivo gravado com sucesso");
			}
			catch (SecurityException ex)
			{
				MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
				$"Details:\n\n{ex.StackTrace}");
			}
		}
	}
}
