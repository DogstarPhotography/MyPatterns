<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="SummaryButton.ascx.vb" Inherits="CSSPatterns.SummaryButton" %>
<style type="text/css">
	#pnlBorder
	{
		border: solid 1px #0000;
		}
	#lblContent
	{
		float: left;
		}
	#btnExpand
	{
		float: right;
		}
</style>
<asp:Panel ID="pnlBorder" runat="server">
	<asp:Label ID="lblContent" runat="server" Text="Content"></asp:Label>
	<asp:Button ID="btnExpand" runat="server" Text="&gt;&gt;" />
</asp:Panel>

