<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


  response.expires=0

  sHTML = "<html> "
  sHTML = sHTML & "<head><title>Cor ESS 4.8 - Loggin service</title></head> "
'  sHTML = sHTML & ofv.FormStyler()
  'sHTML = sHTML & "<body onbeforeunload='beforeunload(event);' bgcolor=" & session("CMFMENU") & " >"
  sHTML = sHTML & "<body  bgcolor=" & session("CMFMENU") & " >"
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




    ' slfun3 = "style='cursor:pointer ;' onclick='document.nulo.submit();' onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000aa'; "" "
   '  xxsaccion = "novedmenu.asp"
   ' xssInput =""' OFV.GenerarInput("boton", "", "submit", "", "1", "it", "")
     xsFormulario = "<form name=nulo id=nulo style='display:none;' action='novedmenu.asp?tab=view4' target=info method=post> "  'OFV.GenForm("nulo", " aa style='display:none;' ", xsaccion, "info", "post", "", xssInput, "si")


     tt=session("ModMain")
     xverasp = UCASE(left(Session("ASP"),4))

     if tt="WK" AND xverasp="NOVE" THEN
        sHTML=sHTML & xsFormulario

     END IF

     wxs = "<script language=JavaScript>" & Chr(13)


    ' wxs = wxs & "window.close();" & Chr(13)

      if tt="WK" AND xverasp="NOVE" THEN
             wxs = wxs & "document.nulo.submit();"
      END IF
      wxs = wxs & "window.open('','_self','');window.close();" & Chr(13)



     wxs = wxs & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & wxs

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

Response.Write sHTML

%>
