<%@ Page Language="VB" AutoEventWireup="true" CodeBehind="HelloWorld.aspx.vb" Inherits="HelloWorldAjaxWCFServiceWebApp.HelloWorld" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
	<title>Untitled Page</title>
	<style type="text/css">
		#TextArea1
		{
			width: 422px;
			height: 315px;
		}
		#txtLong
		{
			height: 265px;
			width: 651px;
		}
	</style>
</head>
<body>
	<form id="form1" runat="server">
	<asp:ScriptManager ID="ScriptManager1" runat="server">
		<Services>
			<asp:ServiceReference Path="HelloWorldService.svc" />
		</Services>
	</asp:ScriptManager>
	<div>
		This is based on the example given here:
		<a href="http://msdn.microsoft.com/en-us/library/bb924552.aspx">
		http://msdn.microsoft.com/en-us/library/bb924552.aspx</a><br />
		This is a web application with an AJAX enabled WCF service (.svc) that is 
		consumed by a page in the same domain. Amount of code is less than for an ajax 
		enabled web site with a web service (.asmx) and configuration is simpler.<br />
		Namespaces need to be &#39;just right&#39; in order for it to work properly.</div>
	</form>
	<p>
		<input id="txtHello" type="text" value="Your Name Here" />
		<input id="btnHello" type="button" value="Hello World" onclick="return btnHello_onclick()" /></p>
	<p>
		<textarea id="txtLong" name="S1" rows="50" cols="50"></textarea>
		<input id="btnLong" type="button" value="Send Long Text"
			onclick="return btnLong_onclick()" /></p>
	<div id="Results"></div>
	<script language="javascript" type="text/javascript">
		// <!CDATA[

		function btnHello_onclick() {
			var txt = document.getElementById('txtHello');
			var service = new HelloWorldAjaxWCFServiceWebApp.HelloWorldService();
			//service.Echo('echo', onSuccess, FailedCallback, null);
			service.HelloWorld(txt.value, onSuccess, FailedCallback, null);
		}

		function btnLong_onclick() {
			var txt = document.getElementById('txtLong');
			var service = new HelloWorldAjaxWCFServiceWebApp.HelloWorldService();
			//service.Echo('echo', onSuccess, FailedCallback, null);
			service.HelloWorld(txt.value, onSuccess, FailedCallback, null);
		}

		function onSuccess(result) {
		    alert(result);
            // This is where we could reload the page (or certain controls) to use any data saved by the service call
		    //window.location.reload(true);
        }

		// This is the failed callback function.
		function FailedCallback(error) {
			var stackTrace = error.get_stackTrace();
			var message = error.get_message();
			var statusCode = error.get_statusCode();
			var exceptionType = error.get_exceptionType();
			var timedout = error.get_timedOut();

			// Display the error.    
			var results = document.getElementById("Results");
			results.innerHTML =
				"Stack Trace: " + stackTrace + "<br/>" +
				"Service Error: " + message + "<br/>" +
				"Status Code: " + statusCode + "<br/>" +
				"Exception Type: " + exceptionType + "<br/>" +
				"Timedout: " + timedout;
		}

		// ]]>
	</script>
</body>
</html>
