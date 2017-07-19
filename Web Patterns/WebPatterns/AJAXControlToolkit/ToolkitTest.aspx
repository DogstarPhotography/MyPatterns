<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ToolkitTest.aspx.vb" Inherits="CSSPatterns.ToolkitTest" %>
<%@ Register TagPrefix="asp" Namespace="AjaxControlToolkit" Assembly="AjaxControlToolkit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
		<ajaxToolkit:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server">
		</ajaxToolkit:ToolkitScriptManager>
		<asp:TextBox ID="TextBox1" runat="server" Height="160px" Width="707px"></asp:TextBox>
		<ajaxToolkit:HtmlEditorExtender
			ID="HtmlEditorExtender1" runat="server" TargetControlID="TextBox1">
		</ajaxToolkit:HtmlEditorExtender>
	</div>
    </form>
</body>
</html>
