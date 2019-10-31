<%@ Language=VBScript %>
<HTML>
	<HEAD>
		<!-- #INCLUDE FILE="Styles.inc" -->
		<!-- #INCLUDE FILE="HyperlinkType.inc" -->
		<!-- #INCLUDE FILE="DataType.inc" -->
		<STYLE> .headerStyle { font-weight:700; background:#FFC020; width:110px}
	.cellStyle { background:#F0F7EF;}
	.footerStyle { background:#F0F7EF;border-bottom:1.0pt solid #FFC020;}
	</STYLE>
		<script Language="JavaScript">
			function exportToExcel() {				
				// Estimated
				if(document.getElementById("chkEstimated").checked)
					document.getElementById("selEstimated").value = "1";
				else
					document.getElementById("selEstimated").value = "0";
					
				// Regular
				if(document.getElementById("chkRegular").checked)
					document.getElementById("selRegular").value = "1";
				else
					document.getElementById("selRegular").value = "0";

				// OTHours
				if(document.getElementById("chkOTHours").checked)
					document.getElementById("selOTHours").value = "1";
				else
					document.getElementById("selOTHours").value = "0";

				// NBHours
				if(document.getElementById("chkNBHours").checked)
					document.getElementById("selNBHours").value = "1";
				else
					document.getElementById("selNBHours").value = "0";

				document.main.action = "EasyXLSDemoApplication.asp?ref="+Math.random();
				document.main.submit();
			}
		</script>
	</HEAD>
	<BODY>
	<TABLE cellspacing="0" cellpadding="0" ID="Table1">
			<TR>
				<TD><img src="EasyXLSlogo.jpg" height="62"></TD>										
			</TR>
			<TR>
				<TD><FONT color="#999999" size="2"><i>* sample image</i></FONT></TD>		
			</TR>
			<TR height=25 valign=bottom>
				<TD><a href="http://www.easyxls.com">www.easyxls.com</a></TD>										
			</TR>
			<TR>
				<TD><FONT color="#999999" size="2"><i>* sample hyperlink</i></FONT></TD>		
			</TR>
	</table>		
		<BR>
		<%

	' Declaring the data
	Dim rows(11)
	rows(0) = Array("Project", "Resource", "Role","Task",  "Estimated", "Regular", "OT Hours", "NB Hours", "Approval Status")
    rows(1) = Array("EasyXLS", "Jim Bean", "Programmer", "Build Charts", 800, 240, 40, 0, "To be Approved")
    rows(2) = Array("EasyXLS", "Jack White", "Programmer", "Build Worksheets", 1000, 160, 0, 0, "To be Approved")
    rows(3) = Array("EasyXLS", "Christina Brown", "Programmer", "Build Hyperlinks", 750, 256, 2, 0, "To be Approved")
    rows(4) = Array("EasyXLS", "Walt Whitman", "Programmer", "Create Tutorials", 600, 114, 10, 0, "To be Approved")
    rows(5) = Array("EasyXLS", "Adam Wilson", "Tester", "Test Charts", 120, 8, 0, 0, "To be Approved")
    rows(6) = Array("EasyXLS", "Will Crane", "Tester", "Test Hyperlinks", 100, 10, 2, 0, "To be Approved")
    rows(7) = Array("EasyXLS", "George Brown", "Artist", "Design", 300, 150, 2, 0, "To be Approved")
    rows(8) = Array("MS Excel", "Christian Wurm", "Programmer", "Database Design", 120, 35, 3, 0, "To be Approved")
    rows(9) = Array("MS Excel", "Adrian Fisher", "Tester", "Speed", 240, 48, 0, 8, "To be Approved")
    
    
    Dim nTotal
    Dim nRow
    Dim nCol
    Dim row
    
    ' Computing the totals
	Dim footerRow
    footerRow = Array("Totals:", "&nbsp;", "&nbsp;", "&nbsp;", "&nbsp;", "&nbsp;", "&nbsp;", "&nbsp;", "&nbsp;")
    nTotal = 0
    For nCol = 4 To 7
        nTotal = 0
        For nRow = 1 To ubound(rows) - 2
			row = rows(nRow)
            nTotal = nTotal + CInt(row(nCol) & "")
        Next
        footerRow(nCol) = nTotal
    Next    		
    rows(10) = footerRow
    
    
    ' Populating the table
    response.Write("<TABLE cellpadding=0 cellspacing=0>")
    For nRow = 0 To ubound(rows)-1  
		response.Write("<TR>") 
		row = rows(nRow)
		For nCol = 0 To ubound(row)
			if nRow = 0  then
				response.Write("<TD CLASS=headerStyle>")
			else
				if nRow = ubound(rows)-1 then
					response.Write("<TD CLASS=footerStyle>")
				else
					response.Write("<TD CLASS=cellStyle>")
				end if
			end if
		    response.Write(row(nCol) & "</TD>")
        Next
        response.Write("</TR>")
    Next
    response.Write("</TABLE>")
 
%>
		<FONT color="#999999" size="2"><i>* sample data set source; totals are computed using 
				formulas</i></FONT>
		<BR>
		<form name="main" action="javascript:exportToExcel();" method="post" ID="Form1">
		<input type="hidden" name="exportToExcel" value="1" ID="Hidden1">
		<input type="hidden" name="selTask" value="1" ID="Hidden2">
		<input type="hidden" name="selEstimated" value="0" ID="Hidden3">
		<input type="hidden" name="selRegular" value="0" ID="Hidden4">
		<input type="hidden" name="selOTHours" value="0" ID="Hidden5">
		<input type="hidden" name="selNBHours" value="0" ID="Hidden6">
		Generate chart with the following columns:&nbsp;

		<TABLE cellspacing="0" cellpadding="0">
			<TR>
				<TD width="25" rowspan="5"></TD>
				<TD><INPUT id="chkTask" type="checkbox" name="chkTask" CHECKED DISABLED></TD>
				<TD>Task</TD>
			</TR>
			<TR>
				<TD><INPUT id="chkEstimated" type="checkbox" name="chkEstimated" CHECKED></TD>
				<TD>Estimated</TD>
			</TR>
			<TR>
				<TD><INPUT id="chkRegular" type="checkbox" name="chkRegular" CHECKED></TD>
				<TD>Regular</TD>
			</TR>
			<TR>
				<TD><INPUT id="chkOTHours" type="checkbox" name="chkOTHours" CHECKED></TD>
				<TD>OT Hours</TD>
			</TR>
			<TR>
				<TD><INPUT id="chkNBHours" type="checkbox" name="chkNBHours" CHECKED></TD>
				<TD>NB Hours</TD>
			</TR>
		</TABLE>
		</form>
		<input type="button" value="Export xls file" onclick="exportToExcel()">		
		<%
    
    Dim exportToExcel
	exportToExcel = request("exportToExcel")
	Dim selTask, selEstimated, selRegular, selOTHours, selNBHours
	selTask = request("selTask")
	selEstimated = request("selEstimated")
	selRegular = request("selRegular")
	selOTHours = request("selOTHours")
	selNBHours = request("selNBHours")
	' Exporting the data to excel
	if ""&exportToExcel = "1" then		
	
		'Create an instance of the object that generates Excel files
		Set xls = Server.CreateObject("EasyXLS.ExcelDocument")
		
		' Adding a sheet to the Excel Document object
		Set xlsWorksheet = Server.CreateObject("EasyXLS.ExcelWorksheet")
		xlsWorksheet.setSheetName("TimeSheetReport")
		xls.easy_addWorksheet(xlsWorksheet)
		
		' Adding the image
		xlsWorksheet.easy_addImage_5 Server.MapPath("EasyXLSlogo.jpg"), "A1"
    
		' Adding the hyperlink
		xlsWorksheet.easy_addHyperlink_2 HYPERLINKTYPE_URL, "http://www.easyxls.com", "A5"


		
		'Create the list used to store the values
		Dim lstRows 
		Set lstRows = CreateObject("EasyXLS.Util.List")
	
		For nRow = 0 To ubound(rows)-1    
			row = rows(nRow)
			Dim lstRow
			Set lstRow = CreateObject("EasyXLS.Util.List")	
			For nCol = 0 To ubound(row)
				lstRow.addElement(row(nCol))
			Next
			lstRows.addElement(lstRow)       	
		Next
		

		'Create an instance of the object used to format the cells
		Dim xlsAutoFormat 
		set xlsAutoFormat = Server.CreateObject("EasyXLS.ExcelAutoFormat")
		xlsAutoFormat.InitAs(AUTOFORMAT_EASYXLS1)

		
		' Adding the data
		xlsWorksheet.easy_insertList_4 lstRows, 6, 0, xlsAutoFormat

		' Creating the footer
		Dim nFooterRowIndex
		nFooterRowIndex = 6 + ubound(rows) - 1
		Dim xlsTable
		Set xlsTable = xlsWorksheet.easy_getExcelTable()
		xlsTable.easy_getCell(nFooterRowIndex, 0).setValue ("Totals:")
		xlsTable.easy_getCell(nFooterRowIndex, 1).setValue ("")
		xlsTable.easy_getCell(nFooterRowIndex, 2).setValue ("")
		xlsTable.easy_getCell(nFooterRowIndex, 3).setValue ("")
		xlsTable.easy_getCell(nFooterRowIndex, 4).setValue ("=SUM(E8:E" & nFooterRowIndex & ")")
		xlsTable.easy_getCell(nFooterRowIndex, 4).setDataType (DATATYPE_AUTOMATIC)
		xlsTable.easy_getCell(nFooterRowIndex, 5).setValue ("=SUM(F8:F" & nFooterRowIndex & ")")
		xlsTable.easy_getCell(nFooterRowIndex, 5).setDataType (DATATYPE_AUTOMATIC)
		xlsTable.easy_getCell(nFooterRowIndex, 6).setValue ("=SUM(G8:G" & nFooterRowIndex & ")")
		xlsTable.easy_getCell(nFooterRowIndex, 6).setDataType (DATATYPE_AUTOMATIC)
		xlsTable.easy_getCell(nFooterRowIndex, 7).setValue ("=SUM(H8:H" & nFooterRowIndex & ")")
		xlsTable.easy_getCell(nFooterRowIndex, 7).setDataType (DATATYPE_AUTOMATIC)
		xlsTable.easy_getCell(nFooterRowIndex, 8).setValue ("")
	       
        ' Creating and adding a chart based on the grid's data	
   		Dim xlsChart 
		Set xlsChart = Server.CreateObject("EasyXLS.Charts.ExcelChart")
		xlsChart.setLeftUpperCorner_2("A20")
		xlsChart.setSize 600, 300

        If ""&selEstimated = "1" Then xlsChart.easy_addSeries_2 "=TimeSheetReport!$E$7", "=TimeSheetReport!$E$8:$E$16"
        If ""&selRegular = "1" Then xlsChart.easy_addSeries_2 "=TimeSheetReport!$F$7", "=TimeSheetReport!$F$8:$F$16"
        If ""&selOTHours = "1" Then xlsChart.easy_addSeries_2"=TimeSheetReport!$G$7", "=TimeSheetReport!$G$8:$G$16"
        If ""&selNBHours = "1" Then xlsChart.easy_addSeries_2"=TimeSheetReport!$H$7", "=TimeSheetReport!$H$8:$H$16"

        If (""&selEstimated = "1" Or ""&selRegular = "1" Or ""&selOTHours = "1" Or ""&selNBHours = "1") Then
            xlsChart.easy_setCategoryXAxisLabels("=TimeSheetReport!$D$8:$D$16")
        Else
            xlsChart.easy_addSeries_2 "=TimeSheetReport!$D$7", "=TimeSheetReport!$D$8:$D$16"
        End If

        xlsWorksheet.easy_addChart(xlsChart)

		'Generate the file
		response.write("Writing file: C:\Samples\ASPApplication.xls<br>")
		xls.easy_WriteXLSFile ("C:\Samples\ASPApplication.xls")
	
		
		'Confirm generation
		if xls.easy_getError() = "" then
			response.write("File successfully created.")
		else
			response.write("Error encountered: " + xls.easy_getError())
		end if
		
	
		'Dispose memory
		xls.Dispose
	end if	

%>
	</BODY>
</HTML>
