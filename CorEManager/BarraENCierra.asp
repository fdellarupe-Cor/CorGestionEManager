<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


  response.expires=0

  sHTML = "<html> "
  sHTML = sHTML & "<head><title>Cor ESS 2.8 - Loggin service</title></head> "
  sHTML = sHTML & ofv.FormStyler()
  sHTML = sHTML & "<body bgcolor=" & session("CMFMENU") & " >"
  sHTML = sHTML & "<p><font color=#993333 size=3 face=tahoma >Actualizando log....</font></p>"

  IF NOT (Session("Usuario") & "") <> "" THEN

         Session("Usuario")     = REPLACE(REPLACE(request.querystring("USR"),"_"," "),"*","")
         Session("UsuarNT")     = REPLACE(REPLACE(request.querystring("NETUSR"),"_"," "),"*","")
         Session("DOMINIO")     = REPLACE(REPLACE(request.querystring("DMN"),"_"," "),"*","")
         SESSION("LOGIDUSRDET") = REPLACE(REPLACE(request.querystring("LID"),"_"," "),"*","")
         SESSION("LOGIDDET")    = REPLACE(REPLACE(request.querystring("LIDOBJ"),"_"," "),"*","")

  END IF

  set cn = ofv.conectar(ofv.strconn2)



     XXXX = GETOBJALOGUPD("FIN")



     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & "window.close(); " & Chr(13)
     wxs = wxs & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & wxs

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

Response.Write sHTML

%>
