<%@ Page Language="VB" AutoEventWireup="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
	<style type="text/css">
            body {  font: 11pt Trebuchet MS;
                    color: #000000;
                    padding-top: 72px;
                    text-align: center }

            .text { font: 8pt Trebuchet MS }
        </style>
	<title>Using AJAX Enabled ASP.NET Service</title>
</head>
<body>
	<form id="Form1" runat="server">
	<asp:ScriptManager runat="server" ID="scriptManager">
		<Services>
			<asp:ServiceReference Path="~/HelloWorld.asmx" />
		</Services>
		<Scripts>
			<asp:ScriptReference Path="~/HelloWorldAjaxProxy.js" />
		</Scripts>
	</asp:ScriptManager>
	<p>
	This project was built following the instructions at
	<a href="http://msdn.microsoft.com/en-us/library/bb532367(v=vs.90).aspx">
	http://msdn.microsoft.com/en-us/library/bb532367(v=vs.90).aspx</a>
	</p>
	<p>
	This is a web service (.asmx) as opposed to a WCF service (.svc) so needs to be part of a website and have the code in the App_Code ASP.NET folder. 
	Javascript in the HelloWorldAjaxProxy is used to pull data from the web service and display it in the web page. 
	Both the web page and the web service need to be in the same site as cross domain calls are prevented by default.
	</p>
	<p>
	The <strong>HelloWorldAjaxWebServiceSite.WebService</strong> Namespace in the App_Code/HelloWorld.vb file is vital to the correct functioning and <strong>
		must</strong> be a two part namespace to work.
	</p>
	<div>
		<h3>
			Using AJAX Enabled ASP.NET Service</h3>
		<table>
			<tr align="left">
				<td>
					Click to call the Hello World service:
				</td>
			</tr>
			<tr align="left">
				<td>
					<button id="Button1" onclick="OnClickGreetings(); return false;">
						Greetings</button>
				</td>
			</tr>
		</table>
	</div>
	</form>
	<hr />
	<div>
		<span id="Results"></span>
	</div>
</body>
</html>
