<%@ Page language="c#" Codebehind="WebForm1.aspx.cs" AutoEventWireup="false" Inherits="CSharpWebApplication.WebForm1" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>WebForm1</title>
		<meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body ms_positioning="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:DataGrid id="dgTimeSheetReport" style="Z-INDEX: 101; POSITION: absolute; TOP: 144px; LEFT: 48px"
				runat="server" Width="800px" ShowFooter="True" DataSource="<%# dsSource %>" AutoGenerateColumns="False" Font-Size="Small">
				<FooterStyle Font-Bold="True" BackColor="LightGreen"></FooterStyle>
				<AlternatingItemStyle BackColor="FloralWhite"></AlternatingItemStyle>
				<ItemStyle BackColor="#F0F7EF"></ItemStyle>
				<HeaderStyle Font-Bold="True" BackColor="LightGreen"></HeaderStyle>
				<Columns>
					<asp:BoundColumn DataField="Project" SortExpression="Project" HeaderText="Project" FooterText="Totals:"></asp:BoundColumn>
					<asp:BoundColumn DataField="Resource" HeaderText="Resource"></asp:BoundColumn>
					<asp:BoundColumn DataField="Role" HeaderText="Role"></asp:BoundColumn>
					<asp:BoundColumn DataField="Task" HeaderText="Task"></asp:BoundColumn>
					<asp:BoundColumn DataField="Estimated" HeaderText="Estimated"></asp:BoundColumn>
					<asp:BoundColumn DataField="Regular" HeaderText="Regular"></asp:BoundColumn>
					<asp:BoundColumn DataField="OT Hours" HeaderText="OT Hours"></asp:BoundColumn>
					<asp:BoundColumn DataField="NB Hours" HeaderText="NB Hours"></asp:BoundColumn>
					<asp:BoundColumn DataField="Approval Status" HeaderText="Approval Status"></asp:BoundColumn>
				</Columns>
			</asp:DataGrid>
			<asp:Label id="Label3" style="Z-INDEX: 111; POSITION: absolute; TOP: 112px; LEFT: 48px" runat="server"
				Font-Italic="True" Font-Size="X-Small" ForeColor="#999999">* sample hyperlink</asp:Label>
			<asp:CheckBox id="chkNBHours" style="Z-INDEX: 109; POSITION: absolute; TOP: 528px; LEFT: 88px"
				runat="server" Text="NB Hours" Checked="True"></asp:CheckBox>
			<asp:CheckBox id="chkOTHours" style="Z-INDEX: 108; POSITION: absolute; TOP: 508px; LEFT: 88px"
				runat="server" Text="OT Hours" Checked="True"></asp:CheckBox>
			<asp:CheckBox id="chkRegular" style="Z-INDEX: 107; POSITION: absolute; TOP: 488px; LEFT: 88px"
				runat="server" Text="Regular" Checked="True"></asp:CheckBox>
			<asp:CheckBox id="chkEstimated" style="Z-INDEX: 106; POSITION: absolute; TOP: 468px; LEFT: 88px"
				runat="server" Text="Estimated" Checked="True"></asp:CheckBox>
			<asp:Button id="btnExportToExcel" style="Z-INDEX: 100; POSITION: absolute; TOP: 560px; LEFT: 48px"
				runat="server" Text="Export xls file"></asp:Button>
			<asp:Image id="imgEasyXLSlogo" style="Z-INDEX: 102; POSITION: absolute; TOP: 8px; LEFT: 40px"
				runat="server" ImageUrl="EasyXLSlogo.jpg"></asp:Image>
			<asp:HyperLink id="hlkEasyXLS" style="Z-INDEX: 103; POSITION: absolute; TOP: 88px; LEFT: 48px"
				runat="server" NavigateUrl="http://www.easyxls.com">www.easyxls.com</asp:HyperLink>
			<asp:Label id="Label1" style="Z-INDEX: 104; POSITION: absolute; TOP: 424px; LEFT: 48px" runat="server"> Generate sheet with the following columns:</asp:Label>
			<asp:CheckBox id="chkTask" style="Z-INDEX: 105; POSITION: absolute; TOP: 448px; LEFT: 88px" runat="server"
				Text="Task" Checked="True" Enabled="False"></asp:CheckBox>
			<asp:Label id="Label2" style="Z-INDEX: 110; POSITION: absolute; TOP: 64px; LEFT: 48px" runat="server"
				Font-Italic="True" Font-Size="X-Small" ForeColor="#999999">* sample image</asp:Label>
			<asp:Label id="Label4" style="Z-INDEX: 113; POSITION: absolute; TOP: 392px; LEFT: 48px" runat="server"
				Font-Italic="True" Font-Size="X-Small" ForeColor="#999999">* sample data set source; totals are computed using formulas</asp:Label>
			<asp:Label id="lblError" style="Z-INDEX: 114; POSITION: absolute; TOP: 560px; LEFT: 184px"
				runat="server" ForeColor="Red"></asp:Label>
		</form>
	</body>
</HTML>
