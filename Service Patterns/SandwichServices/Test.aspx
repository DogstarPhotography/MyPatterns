<%@ Page Language="C#" %>

<%@ Import Namespace="System.Runtime.Remoting" %>
<%
ObjectHandle oh = Activator.CreateInstance("SandwichServices", "SandwichServices.CostService");
Response.Write(oh.Unwrap().ToString() + " loaded OK");
%>
