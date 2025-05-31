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
			Filter = "Arquivos Excel (*.xlsx)|*.xlsx",
			Multiselect = true
		};


		if (_openFileDialog.ShowDialog() == DialogResult.OK)
		{
			try
			{
				var fileName = _openFileDialog.FileName;

				var newFilePath = string.Empty;

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

						var directory = Path.GetDirectoryName(fileName);

						var csvFilePath = Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(fileName)}.csv");

						newFilePath = csvFilePath.ToString();

						using var file = File.OpenWrite(csvFilePath);

						using var streamWriter = new StreamWriter(file);

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

							streamWriter.WriteLine(line);
							streamWriter.WriteLine();
						}
					}
				}

				var newCsvDirectory = Path.GetDirectoryName(newFilePath);

				var xlsxFilePath = Path.Combine(newCsvDirectory, $"{Path.GetFileNameWithoutExtension(fileName)} - COM LINHAS.xlsx");

				using (SpreadsheetDocument document = SpreadsheetDocument.Create(xlsxFilePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
				{
					WorkbookPart workbookPart = document.AddWorkbookPart();
					workbookPart.Workbook = new Workbook();

					WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
					SheetData sheetData = new SheetData();

					worksheetPart.Worksheet = new Worksheet(sheetData);
					Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
					Sheet sheet = new Sheet()
					{
						Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
						SheetId = 1,
						Name = "Sheet1"
					};
					sheets.Append(sheet);

					uint rowIndex = 1;
					foreach (var line in File.ReadLines(newFilePath))
					{
						// Add an empty row if the line is blank
						if (string.IsNullOrWhiteSpace(line))
						{
							sheetData.Append(new Row() { RowIndex = rowIndex++ });
							continue;
						}

						Row row = new Row() { RowIndex = rowIndex++ };
						string[] values = line.Split(',');

						foreach (var value in values)
						{
							Cell cell = new Cell()
							{
								DataType = CellValues.String,
								CellValue = new CellValue(value)
							};
							row.Append(cell);
						}

						sheetData.Append(row);
					}

					workbookPart.Workbook.Save();
				}

				File.Delete(newFilePath);

				MessageBox.Show($"Arquivo gravado com sucesso");
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Houve um erro: {ex.Message}", "Erro");
			}
		}
	}
}
