<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TestControl.aspx.vb" Inherits="CSSPatterns.TestControl" %>

<%@ Register src="SummaryButton.ascx" tagname="SummaryButton" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    	<uc1:SummaryButton ID="SummaryButton1" runat="server" Text="Lorem Ipsum" />
    
    </div>
    </form>
</body>
</html>
