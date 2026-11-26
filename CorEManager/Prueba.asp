
<%

'pepe = clng(date()) & " - " & replace(time(),":","")  & " - " & time()

pepe = clng(date()) & replace(time(),":","")

response.write pepe

%>
