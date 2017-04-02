using System;
using Excel=Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Drawing;
using System.Collections.Generic;
namespace Allocator
{
    public partial class MainForm : Form
    {
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook workbook;
        string masterWorkbook;
        string masterText = "Open the Task Workbook containing the details of the tasks assigned..";
        int namesIndex = 0;
        int masterColCount = 0;
        int formIdIndex = 0;
        string filePath = null;
        string masterSheetName;

        public MainForm()
        {
            InitializeComponent();
        }

        private void masterButton_Click(object sender, EventArgs e)
        {
            masterWorkbook = null;
            filePath = null;
            masterColCount = 0;
            namesIndex = 0;
            formIdIndex = 0;
            OpenFileDialog opMaster = new OpenFileDialog();
            opMaster.Multiselect = false;
            opMaster.ShowDialog();
            if (opMaster.FileNames.Length == 0)
                MessageBox.Show("No Workbook was selected.\nPlease try again...");
            else
            {
                masterTextBox.Text = "The selected workbook is being analysed. Please wait...";
                if (opMaster.FileName.Substring(opMaster.FileName.Length - 5) != ".xlsx")
                {
                    MessageBox.Show("The selected Workbook is not an Excel Workbook!!! Please try again...");
                    masterTextBox.Text = masterText;
                }
                else
                {
                    try
                    {
                        masterColCount = 0;
                        Excel.Range range;
                        workbook = xlApp.Workbooks.Open(opMaster.FileName);
                        Excel.Worksheet worksheet = workbook.ActiveSheet;
                        int col = 1;
                        while (true)
                        {
                            range = worksheet.Cells[1, col];
                            if (range.Value != null)
                            {
                                masterColCount++;
                                col++;
                                continue;
                            }
                            else
                                break;
                        }

                        while ((namesIndex <= 0 || namesIndex > masterColCount) || (formIdIndex <= 0 || formIdIndex > masterColCount))
                        {
                            try
                            {
                                namesIndex = int.Parse(Showdialog());
                            }
                            catch (FormatException ex)
                            {
                                MessageBox.Show("Enter a valid Integer...");
                            }
                        }
                        masterSheetName = worksheet.Name;
                        masterTextBox.Text = masterWorkbook = opMaster.FileName;
                        workbook.Close();
                        xlApp.Quit();
                        filePath = masterWorkbook.Substring(0, masterWorkbook.LastIndexOf("\\") + 1);
                        MessageBox.Show("The workbook has been successfully added...");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("The selected Workbook cannot be opened!!!Please try again...");
                        masterTextBox.Text = masterText;
                        xlApp.Quit();
                    }
                }
            }
        }

        private void generatorButton_Click(object sender, EventArgs e)
        {
            if (masterWorkbook == null)
            {
                MessageBox.Show("Select the Task Workbook first to proceed...");
            }
            else
            {
                DataSet ds = new DataSet();
                DataTable workbookTable = new DataTable();
                OleDbDataAdapter xlAdapter = new OleDbDataAdapter();
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + masterWorkbook + ";Extended Properties=\"Excel 12.0;HDR=YES;\"";

                OleDbCommand comm = new OleDbCommand();
                comm.Connection = conn;
                conn.Open();
                workbookTable = comm.Connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                comm.CommandText = "SELECT * FROM [" + workbookTable.Rows[0]["TABLE_NAME"].ToString() + "]";
                xlAdapter.SelectCommand = comm;
                xlAdapter.Fill(ds);
                conn.Close();
                IEnumerable<string> namesCollection = (from row in ds.Tables[0].AsEnumerable()
                                                       select row[namesIndex - 1].ToString()).ToList().Distinct();

                foreach (string name in namesCollection)
                {
                    var table = from row in ds.Tables[0].AsEnumerable()
                                where row[namesIndex - 1].ToString() == name
                                select row;

                    workbook = xlApp.Workbooks.Add();
                    Excel.Worksheet worksheet = workbook.ActiveSheet;               
                    try
                    {
                        workbook.SaveAs(filePath + name + ".xlsx");
                        
                    }
                    catch (Exception ex)
                    {
                        workbook.Close();
                        continue;
                    }

                    worksheet.Name = masterSheetName;
                    for (int i = 1; i <= masterColCount; i++)
                    {
                        worksheet.Cells[1, i] = ds.Tables[0].Columns[i - 1].ColumnName.ToString();
                    }
                    workbook.Save();
                    workbook.Close();
                    conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + name + ".xlsx;Extended Properties=\"Excel 12.0;HDR=YES;\"";
                    comm.Connection = conn;
                    conn.Open();
                    workbookTable.Clear();
                    workbookTable = comm.Connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                
                    foreach (DataRow row in table)
                    {
                        comm.CommandText = "INSERT INTO [" + masterSheetName + "$] VALUES(";
                        for (int i = 0; i < masterColCount; i++)
                        {
                            if (i == masterColCount - 1)
                                comm.CommandText += "'" + row[i].ToString() + "')";
                            else
                                comm.CommandText += "'" + row[i].ToString() + "',";
                        }
                        comm.ExecuteNonQuery();
                    }
                    conn.Close();
                    xlApp.Quit();
                }
                MessageBox.Show(string.Format("The individual excel files have been generated and placed at the path:\n{0}", filePath));
            }
        }

        private string Showdialog()
        {
            Form indexForm = new Form();
            indexForm.Size = new System.Drawing.Size(300, 180);
            indexForm.MinimizeBox = false;
            indexForm.MaximizeBox = false;
            indexForm.StartPosition = FormStartPosition.CenterParent;
            indexForm.FormBorderStyle = FormBorderStyle.FixedSingle;
            indexForm.Text = "Enter Master Column Index";

            Label label1 = new Label();
            label1.Text = "Names Column Index starting from 1:";
            label1.Top = 20;
            label1.Left = 17;
            label1.Width = 250;
            label1.Height = 30;
            label1.BorderStyle = BorderStyle.Fixed3D;
            label1.TextAlign = ContentAlignment.MiddleLeft;

            TextBox namesIndexText = new TextBox();
            namesIndexText.Width = 50;
            namesIndexText.Top = 4;
            namesIndexText.Left = 183;
            indexForm.Controls.Add(namesIndexText);
            label1.Controls.Add(namesIndexText);
            indexForm.Controls.Add(label1);

            Label label2 = new Label();
            label2.Text = "FormID Column Index starting from 1:";
            label2.Top = 70;
            label2.Left = 17;
            label2.Width = 250;
            label2.Height = 30;
            label2.BorderStyle = BorderStyle.Fixed3D;
            label2.TextAlign = ContentAlignment.MiddleLeft;

            TextBox formIdIndexText = new TextBox();
            formIdIndexText.Width = 50;
            formIdIndexText.Top = 4;
            formIdIndexText.Left = 183;
            indexForm.Controls.Add(formIdIndexText);
            label2.Controls.Add(formIdIndexText);
            indexForm.Controls.Add(label2);

            Button acceptButton = new Button();
            acceptButton.Text = "OK";
            acceptButton.Left = 110;
            acceptButton.Top = 110;
            acceptButton.DialogResult = DialogResult.OK;
            indexForm.Controls.Add(acceptButton);
            indexForm.AcceptButton = acceptButton;
           
            if (indexForm.ShowDialog(this) == DialogResult.OK)
            {
                namesIndex = int.Parse(namesIndexText.Text);
                formIdIndex = int.Parse(formIdIndexText.Text);
                return namesIndexText.Text;
            }
            else return null;
        }
    }
}