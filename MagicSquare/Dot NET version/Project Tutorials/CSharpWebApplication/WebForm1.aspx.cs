using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;

using EasyXLS;
using EasyXLS.Charts;
using EasyXLS.Constants;


namespace CSharpWebApplication
{
	/// <summary>
	/// Summary description for WebForm1.
	/// </summary>
	public class WebForm1 : System.Web.UI.Page
	{
		protected System.Data.DataSet dsSource;
		protected System.Data.DataColumn dataColumn1;
		protected System.Data.DataColumn dataColumn2;
		protected System.Data.DataColumn dataColumn3;
		protected System.Data.DataColumn dataColumn4;
		protected System.Data.DataColumn dataColumn5;
		protected System.Data.DataColumn dataColumn6;
		protected System.Data.DataColumn dataColumn7;
		protected System.Data.DataColumn dataColumn8;
		protected System.Data.DataColumn dataColumn9;
		protected System.Data.DataTable dtTable;
		protected System.Web.UI.WebControls.DataGrid dgTimeSheetReport;
		
		protected System.Web.UI.WebControls.Button btnExportToExcel;
		protected System.Web.UI.WebControls.HyperLink hlkEasyXLS;
		protected System.Web.UI.WebControls.Image imgEasyXLSlogo;		
		protected System.Web.UI.WebControls.Label Label1;
		protected System.Web.UI.WebControls.CheckBox chkTask;
		protected System.Web.UI.WebControls.CheckBox chkEstimated;
		protected System.Web.UI.WebControls.CheckBox chkRegular;
		protected System.Web.UI.WebControls.CheckBox chkOTHours;
		protected System.Web.UI.WebControls.CheckBox chkNBHours;
		protected System.Web.UI.WebControls.Label Label2;
		protected System.Web.UI.WebControls.Label Label3;
		protected System.Web.UI.WebControls.Label Label4;
		protected System.Web.UI.WebControls.Label lblError;


		private void Page_Load(object sender, System.EventArgs e)
		{
			// Populating the grid
			dtTable.Rows.Add(new Object[] {"EasyXLS", "Jim Bean", "Programmer", "Build Charts", 800, 240, 40, 0, "To be Approved"});
			dtTable.Rows.Add(new Object[] {"EasyXLS", "Jack White", "Programmer", "Build Worksheets", 1000, 160, 0, 0, "To be Approved"});
			dtTable.Rows.Add(new Object[] {"EasyXLS", "Christina Brown", "Programmer", "Build Hyperlinks", 750, 256, 2, 0, "To be Approved"});
			dtTable.Rows.Add(new Object[] {"EasyXLS", "Walt Whitman", "Programmer", "Create Tutorials", 600, 114, 10, 0, "To be Approved"});
			dtTable.Rows.Add(new Object[] {"EasyXLS", "Adam Wilson", "Tester", "Test Charts", 120, 8, 0, 0, "To be Approved"});
			dtTable.Rows.Add(new Object[] {"EasyXLS", "Will Crane", "Tester", "Test Hyperlinks", 100, 10, 2, 0, "To be Approved"});
			dtTable.Rows.Add(new Object[] {"EasyXLS", "George Brown", "Artist", "Design", 300, 150, 2, 0, "To be Approved"});
			dtTable.Rows.Add(new Object[] {"MS Excel", "Christian Wurm", "Programmer", "Database Design", 120, 35, 3, 0, "To be Approved"});
			dtTable.Rows.Add(new Object[] {"MS Excel", "Adrian Fisher", "Tester", "Speed", 240, 48, 0, 8, "To be Approved"});

			// Computing the totals
			int nTotal = 0;
			for (int nColumnIndex = 4; nColumnIndex < 8; nColumnIndex++)
			{
				nTotal = 0;
				for (int nRowIndex = 0; nRowIndex < dtTable.Rows.Count; nRowIndex++)
				{
					nTotal += int.Parse(dtTable.Rows[nRowIndex].ItemArray[nColumnIndex].ToString());
				}
				dgTimeSheetReport.Columns[nColumnIndex].FooterText = nTotal + "";
			}

			// Data binding
			dgTimeSheetReport.DataBind();				
		}

		

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    
			this.dsSource = new System.Data.DataSet();
			this.dtTable = new System.Data.DataTable();
			this.dataColumn1 = new System.Data.DataColumn();
			this.dataColumn2 = new System.Data.DataColumn();
			this.dataColumn3 = new System.Data.DataColumn();
			this.dataColumn4 = new System.Data.DataColumn();
			this.dataColumn5 = new System.Data.DataColumn();
			this.dataColumn6 = new System.Data.DataColumn();
			this.dataColumn7 = new System.Data.DataColumn();
			this.dataColumn8 = new System.Data.DataColumn();
			this.dataColumn9 = new System.Data.DataColumn();
			((System.ComponentModel.ISupportInitialize)(this.dsSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtTable)).BeginInit();
			this.btnExportToExcel.Click += new System.EventHandler(this.btnExportToExcel_Click);
			// 
			// dsSource
			// 
			this.dsSource.DataSetName = "dsSource";
			this.dsSource.Locale = new System.Globalization.CultureInfo("en-US");
			this.dsSource.Tables.AddRange(new System.Data.DataTable[] {
																		  this.dtTable});
			// 
			// dtTable
			// 
			this.dtTable.Columns.AddRange(new System.Data.DataColumn[] {
																		   this.dataColumn1,
																		   this.dataColumn2,
																		   this.dataColumn3,
																		   this.dataColumn4,
																		   this.dataColumn5,
																		   this.dataColumn6,
																		   this.dataColumn7,
																		   this.dataColumn8,
																		   this.dataColumn9});
			this.dtTable.TableName = "dtTable";
			// 
			// dataColumn1
			// 
			this.dataColumn1.ColumnName = "Project";
			// 
			// dataColumn2
			// 
			this.dataColumn2.ColumnName = "Resource";
			// 
			// dataColumn3
			// 
			this.dataColumn3.ColumnName = "Role";
			// 
			// dataColumn4
			// 
			this.dataColumn4.ColumnName = "Task";
			// 
			// dataColumn5
			// 
			this.dataColumn5.ColumnName = "Estimated";
			this.dataColumn5.DataType = typeof(int);
			// 
			// dataColumn6
			// 
			this.dataColumn6.ColumnName = "Regular";
			this.dataColumn6.DataType = typeof(int);
			// 
			// dataColumn7
			// 
			this.dataColumn7.ColumnName = "OT Hours";
			this.dataColumn7.DataType = typeof(int);
			// 
			// dataColumn8
			// 
			this.dataColumn8.ColumnName = "NB Hours";
			this.dataColumn8.DataType = typeof(int);
			// 
			// dataColumn9
			// 
			this.dataColumn9.ColumnName = "Approval Status";
			this.Load += new System.EventHandler(this.Page_Load);
			((System.ComponentModel.ISupportInitialize)(this.dsSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtTable)).EndInit();

		}
		#endregion

		private void btnExportToExcel_Click(object sender, System.EventArgs e)
		{
			// Creating an instance of the object that generates excel files
			ExcelDocument xls = new ExcelDocument();
			
			// Adding a sheet to the Excel Document object
			ExcelWorksheet xlsWorksheet = new ExcelWorksheet("TimeSheetReport");
			xls.easy_addWorksheet(xlsWorksheet);

			// Adding the image
			xlsWorksheet.easy_addImage(Server.MapPath( imgEasyXLSlogo.ImageUrl), "A1"); 

			// Adding the hyperlink
			xlsWorksheet.easy_addHyperlink(HyperlinkType.URL, hlkEasyXLS.NavigateUrl, "A5");


			// Creating an instance of the object used to format the cells
			ExcelAutoFormat xlsAutoFormat = new ExcelAutoFormat();
			// Setting the style of the header
			ExcelStyle xlsHeaderStyle = new ExcelStyle(dgTimeSheetReport.HeaderStyle.BackColor);
			xlsHeaderStyle.setBold(dgTimeSheetReport.HeaderStyle.Font.Bold);			
			xlsAutoFormat.setHeaderRowStyle(xlsHeaderStyle);
			// Setting the style of the cells
			xlsAutoFormat.setEvenRowStripesStyle(new ExcelStyle(dgTimeSheetReport.ItemStyle.BackColor));
			xlsAutoFormat.setOddRowStripesStyle(new ExcelStyle(dgTimeSheetReport.AlternatingItemStyle.BackColor));
			
			// Adding the content of the grid
			xlsWorksheet.easy_insertDataSet(dsSource, 6, 0, xlsAutoFormat, true);

			// Creating the footer
			int nFooterRowIndex = 6 + dtTable.Rows.Count + 1;
			ExcelTable xlsTable = xlsWorksheet.easy_getExcelTable();
			xlsTable.easy_getCell(nFooterRowIndex, 0).setValue("Totals:");
			xlsTable.easy_getCell(nFooterRowIndex, 4).setValue("=SUM(E8:E" + nFooterRowIndex + ")");
			xlsTable.easy_getCell(nFooterRowIndex, 5).setValue("=SUM(F8:F" + nFooterRowIndex + ")");
			xlsTable.easy_getCell(nFooterRowIndex, 6).setValue("=SUM(G8:G" + nFooterRowIndex + ")");
			xlsTable.easy_getCell(nFooterRowIndex, 7).setValue("=SUM(H8:H" + nFooterRowIndex + ")");
			// Setting the style of the footer
			ExcelStyle xlsFooterStyle = new ExcelStyle(dgTimeSheetReport.FooterStyle.BackColor);
			xlsFooterStyle.setBold(dgTimeSheetReport.FooterStyle.Font.Bold);
			xlsTable.easy_setRangeStyle(nFooterRowIndex, 0, nFooterRowIndex, 8, xlsFooterStyle);

			
			// Creating and adding a chart based on the grid's data	
			ExcelChart xlsChart = new ExcelChart("A20", 600, 300);	
			if (chkEstimated.Checked)
				xlsChart.easy_addSeries("=TimeSheetReport!$E$7", "=TimeSheetReport!$E$8:$E$16");
			if (chkRegular.Checked)
				xlsChart.easy_addSeries("=TimeSheetReport!$F$7", "=TimeSheetReport!$F$8:$F$16");
			if (chkOTHours.Checked)
				xlsChart.easy_addSeries("=TimeSheetReport!$G$7", "=TimeSheetReport!$G$8:$G$16");
			if (chkNBHours.Checked)
				xlsChart.easy_addSeries("=TimeSheetReport!$H$7", "=TimeSheetReport!$H$8:$H$16");

			if (chkEstimated.Checked || chkRegular.Checked || chkOTHours.Checked || chkNBHours.Checked)
				xlsChart.easy_setCategoryXAxisLabels("=TimeSheetReport!$D$8:$D$16");
			else
				xlsChart.easy_addSeries("=TimeSheetReport!$D$7", "=TimeSheetReport!$D$8:$D$16");

			xlsWorksheet.easy_addChart(xlsChart);

			// Preparing the Response object
			Response.AppendHeader("content-disposition", "attachment; filename=CSharpWebApplication.xls");
			Response.ContentType = "application/octetstream";
			Response.Clear();

			// Generating the file and prompting the "Open or Save Dialog Box"
			try
			{
				xls.easy_WriteXLSFile(Response.OutputStream);
			}
			catch (Exception exc)
			{
				Response.ClearHeaders();
				Response.ClearContent();
				lblError.Text = exc.Message;
			}
			xls.Dispose();
		}
	}
}
