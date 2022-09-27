using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using ClosedXML.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Demo_BI_Gen_Data_Tool
{
    public partial class Input_File : Form
    {
        public Input_File()
        {
            InitializeComponent();
        }
        DataTableCollection dataTableCollection;

        public void getData(string filename)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename);
            Excel.Worksheet sheet = xlWorkbook.Sheets["DIM_Calendar"];
            Console.WriteLine(xlWorkbook);
            Console.WriteLine(sheet);
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            using(OpenFileDialog dlg = new OpenFileDialog() 
            { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls" })
            {
                if(dlg.ShowDialog() == DialogResult.OK)
                {
                    txtFileName.Text = dlg.FileName;
                    /*using(var stream = File.Open(dlg.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using(IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration() 
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true}
                            });
                            dataTableCollection = result.Tables;
                            cbxData.DataSource = null;
                            cbxSheet.Items.Clear();
                            cbxSheet.Items.AddRange(dataTableCollection.Cast<DataTable>().Select(t => t.TableName).ToArray<string>());
                        }
                    }*/
                    Console.WriteLine(txtFileName.Text);
                    getData(dlg.FileName.ToString());
                }
                
            }
        }

        DataTable dt;
        private void cbxSheet_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //get column by sheet
            dt = dataTableCollection[cbxSheet.SelectedItem.ToString()];
            var columnName = (from c in dt.Columns.Cast<DataColumn>()
                              select c.ColumnName).ToArray();
            cbxColumn.Items.Clear();
            cbxColumn.Items.AddRange(columnName);
            Console.WriteLine(dt);

            

            //Display on grid view
            dataGrid.DataSource = dt;
            dataGrid.AutoResizeColumns();

            for (int i = 1; i < dataGrid.RowCount - 1; i++)
            {
                if (dataGrid.Rows[i].Cells[0].Value.ToString() == "" || dataGrid.Rows[i].Cells[1].Value.ToString() == "")
                {
                    dataGrid.Rows.RemoveAt(i);
                    i--;

                }
            }

            //Count row
            lblCount.Text = (dataGrid.RowCount - 1).ToString();
        }

        private void cbxColumn_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //get data by column
            if(dt != null)
            {
                string columnName = cbxColumn.SelectedItem.ToString();
                var data = dt.DefaultView.ToTable(true, columnName);
                cbxData.DataSource = data;
                cbxData.DisplayMember = columnName;
                cbxData.ValueMember = columnName;

                dataGrid.DataSource = data;
                dataGrid.AutoResizeColumns();
            }
            //Count row
            lblCount.Text = (dataGrid.RowCount - 1).ToString();
        }

        private void cbxData_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (dt != null)
            {
                string columnName = cbxColumn.SelectedItem.ToString();
                string valueName = cbxData.SelectedValue.ToString();

                var data = dt.DefaultView.ToTable(true, columnName);
                dataGrid.AutoResizeColumns();
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //dt = dataTableCollection[cbxSheet.SelectedItem.ToString()];
            dataGrid.DataSource = dt;
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter= "Excel Workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using(XLWorkbook workbook = new XLWorkbook())
                        {
                            workbook.Worksheets.Add(dt);
                            workbook.SaveAs(sfd.FileName);
                        }
                        MessageBox.Show("ok", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void Input_File_Load(object sender, EventArgs e)
        {
            dateFrom.Value = new DateTime(2022, 09, 30);
            /*getData();*/
        }
    }
}
