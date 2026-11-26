<%
Response.ContentType = "application/msword"
Response.AddHeader "Content-Disposition", "attachment;filename=receipt.doc"

Dim strName, strAddress

strName = "2222"
strAddress = "xxxxxx"

%>
<html>

<head>
<title>Word Test</title>
<meta name="Author" content="Acme Toys">

<style type="text/css">

.BigTitle {
font-family: Verdana, Sans-Serif;
font-size: 20pt;
font-weight: bold;
color: #004080;
}

.UserDetails {
font-family: Courier New, Monospace;
font-size: 12pt;
}

</style>

</head>

<body>

<span class="BigTitle">Sample Receipt</span>

<p>
Thank you for purchasing, your order will be shipped within 2 business
days. Below are the shipping details, provided by you.
</p>

<p>
<table cellspacing="1" cellpadding="1" border="0">
<tr>
<td><b>Name:</b></td>
<td><span class="UserDetails"><%=strName%></span></td>
</tr>
<tr>
<td><b>Address:</b></td>
<td><span class="UserDetails"><%=strAddress%></span></td>
</tr>
</table>
</p>

<table cellspacing="1" cellpadding="1" border="0">

<%
 For Each i in Session.Contents

 response.write "<tr><td><b>"& i &":</b></td>"
 response.write "<td><span class=UserDetails>"&SESSION(I) &"</span></td></tr>"
 next
 %>



</table>

<table cellspacing="1" cellpadding="1" border="0">

<%
 For Each i in Request.ServerVariables

 response.write "<tr><td><b>"& i &":</b></td>"
 response.write "<td><span class=UserDetails>"&Request.ServerVariables(I) &"</span></td></tr>"
 next
 %>



</table>


</body>

</html>