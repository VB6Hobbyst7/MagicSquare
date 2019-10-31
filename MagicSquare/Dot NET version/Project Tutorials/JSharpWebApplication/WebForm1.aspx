<%@ Page language="VJ#" Codebehind="WebForm1.aspx.jsl" AutoEventWireup="false" Inherits="JSharpWebApplication.WebForm1" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>WebForm1</title>
		<meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" Content="VJ#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:Button id="btnExportToExcel" style="Z-INDEX: 105; LEFT: 48px; POSITION: absolute; TOP: 560px"
				runat="server" Text="Export xls file"></asp:Button>
			<asp:Label id="lblError" style="Z-INDEX: 114; LEFT: 192px; POSITION: absolute; TOP: 560px"
				runat="server" ForeColor="Red"></asp:Label>
			<asp:CheckBox id="chkNBHours" style="Z-INDEX: 113; LEFT: 88px; POSITION: absolute; TOP: 528px"
				runat="server" Text="NB Hours" Checked="True"></asp:CheckBox>
			<asp:CheckBox id="chkOTHours" style="Z-INDEX: 108; LEFT: 88px; POSITION: absolute; TOP: 508px"
				runat="server" Text="OT Hours" Checked="True"></asp:CheckBox>
			<asp:CheckBox id="chkRegular" style="Z-INDEX: 107; LEFT: 88px; POSITION: absolute; TOP: 488px"
				runat="server" Text="Regular" Checked="True"></asp:CheckBox>
			<asp:CheckBox id="chkEstimated" style="Z-INDEX: 106; LEFT: 88px; POSITION: absolute; TOP: 468px"
				runat="server" Text="Estimated" Checked="True"></asp:CheckBox>
			<asp:Label id="Label1" style="Z-INDEX: 103; LEFT: 48px; POSITION: absolute; TOP: 424px" runat="server"> Generate chart with the following columns:</asp:Label>
			<asp:CheckBox id="chkTask" style="Z-INDEX: 104; LEFT: 88px; POSITION: absolute; TOP: 448px" runat="server"
				Text="Task" Checked="True" Enabled="False"></asp:CheckBox>
			<asp:Label id="Label4" style="Z-INDEX: 112; LEFT: 48px; POSITION: absolute; TOP: 392px" runat="server"
				Font-Italic="True" ForeColor="#999999" Font-Size="X-Small">* sample data set source; totals are computed using formulas</asp:Label>
			<asp:Label id="Label3" style="Z-INDEX: 110; LEFT: 48px; POSITION: absolute; TOP: 112px" runat="server"
				Font-Italic="True" ForeColor="#999999" Font-Size="X-Small">* sample hyperlink</asp:Label>
			<asp:Label id="Label2" style="Z-INDEX: 109; LEFT: 48px; POSITION: absolute; TOP: 64px" runat="server"
				Font-Italic="True" ForeColor="#999999" Font-Size="X-Small">* sample image</asp:Label>
			<asp:Image id="imgEasyXLSlogo" style="Z-INDEX: 101; LEFT: 40px; POSITION: absolute; TOP: 8px"
				runat="server" ImageUrl="EasyXLSlogo.jpg"></asp:Image>
			<asp:HyperLink id="hlkEasyXLS" style="Z-INDEX: 102; LEFT: 48px; POSITION: absolute; TOP: 88px"
				runat="server" NavigateUrl="http://www.easyxls.com">www.easyxls.com</asp:HyperLink>
			<asp:DataGrid id=dgTimeSheetReport style="Z-INDEX: 100; LEFT: 48px; POSITION: absolute; TOP: 144px" runat="server" Width="800px" ShowFooter="True" DataSource="<%# dsSource %>" AutoGenerateColumns="False">
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
		</form>
	</body>
</HTML>
