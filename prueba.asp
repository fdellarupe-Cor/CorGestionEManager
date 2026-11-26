
<html>
    <head>
    <title>

    </title>
    </head>


    <body>


    <%
        Dim objHTTP, strEnvelope
        Set objHTTP = Server.CreateObject("Microsoft.XMLHTTP")

        'Create the SOAP Envelope.
        'Start with standard xml name space and XML Schema Definition.
        strEnvelope = "<?xml version='1.0' encoding='utf-8'?>"
        strEnvelope = strEnvelope & "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"

        'Define body of SOAP with method name and parameter names and values to be passed.

        strEnvelope = strEnvelope & "<soap:Body>"
        strEnvelope = strEnvelope & "<RetornaTexto2 xmlns='http://tempuri.org/'>"
		strEnvelope = strEnvelope & "<texto>Texto de salidaqqqqqqqqqqqqqqqqqqqqq</texto>"
		strEnvelope = strEnvelope & "</RetornaTexto2>"
        strEnvelope = strEnvelope & "</soap:Body></soap:Envelope>"

        Dim url
        url = "http://server-40:8057/Service1.asmx"
        With objHTTP
            .Open "post", url, False
            '.setRequestHeader "Host", "http:///server-40:8057"
            .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
            '.setRequestHeader "SOAPAction", "http://server-40:8057/"
            .send strEnvelope
        End With
        ' Following will write xml received from web services in the browser
        Dim strResponse
        set strResponse = objHTTP.responseXML

        'response.write objHTTP.responseXML.xml

        set x=strResponse.SelectSingleNode("//RetornaTexto2Result")
        response.write x.text


      '  response.write  strResponse
      '  If (strResponse = "") Then
      '      Response.Write("Invalid user")
      '  Else
      '      Set myXmlDoc = Server.CreateObject("MSXML2.DOMDocument")
'
     '       myXmlDoc.loadXML (strResponse)
'
      '      Set objLst = myXmlDoc.getElementsByTagName("retornatextoresponse").item(0).text 'SelectSingleNode("/soap/retornatextoresponse")
'
      '      Set objListNodes = objLst.Context.childNodes(0).childNodes
      '      For i = 0 To (objListNodes.Length - 1)
      '         Response.Write(objListNodes.Item(i).nodeName & ":------ " & objListNodes.Item(i).Text)
      '         Response.Write("</BR>")
      '      Next
      '  End If

    %>


    </body>
</html>