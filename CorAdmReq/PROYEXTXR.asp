<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

'  IF session("ModMain") <> "WF" THEN
'     bAgregar = FALSE
'     bModificar = FALSE
'     bSuprimir = FALSE
'  END IF

 sopcion = request.querystring("opcion")


   sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>&#0128;</FONT></B>"
    sImagenA = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"

    sImagenE0 = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=4>6</FONT></B>"
    sImagenE1 = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"
    sImagenE2 = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=4>&#0251;</FONT></B>"




  StrSql = "Select * From PROYETPTSK where codigo = '" & session("NCliente") & "' AND VPLAN = " & SESSION("VPLAN") & " ORDER BY USRRESP,ORDEN,TSKORD "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then


 

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","cc colspan=3","","","","","",sImagen & "&nbsp;<BIG><B>TAREAS POR RESPONSABLE</B></BIG>","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","vt  ","","","","","","RESPONSABLE","si")
  celdas = celdas & ofv.GenCelda("","vt  ","","","","","","&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","vt  ","","","","","","TAREAS","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  USRANT = "XX"
  usrtsk = 0

  DO UNTIL RS.EOF

     IF USRANT <> RS(13) THEN

      if usrtsk > 0 then
         rowspan= " rowspan=" & (usrtsk + 1) & " "
      celdas = ofv.GenCelda("","vt" & rowspan,"","","","","",USRDESC,"si")
      celdas = CELDAS & ofv.GenCelda("","vt COLSPAN=2 ","","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
 
      filas = filas & FILAS1

      celdas = ofv.GenCelda("","AA COLSPAN=3 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
      end if

    if RS(13) = "NO" then
      sInput = "SIN ASIGNAR"

    else
      strsql = "Select * From USUARIOS where ID = '" & rs(13) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
        sInput = sImagenR & "&nbsp;<B>" & sniv & "</B>"
    END IF 

        USRDESC = sInput
        USRANT = RS(13)
        usrtsk = 0

        FILAS1 = ""
     END IF

        usrtsk =  usrtsk + 1

    select case rs(4)
           case 0
             sImagen = sImagenE0 & "&nbsp;"

           case 1
             sImagen = sImagenE1 

           case 2
             sImagen = sImagenE2 & "&nbsp;"


    end select



     SINPUT = sImagen & "&nbsp;&nbsp;<B>" & RS(0) & "." & RS(1) & "&nbsp;" & RS(8) & "</b>"
     celdas = ofv.GenCelda("","vt COLSPAN=2 ","","","","","",sInput,"si")


      filas1 = filas1 & ofv.GenRow("","","","","10","","",celdas,"si")




    RS.MOVENEXT
    LOOP

      if usrtsk > 0 then
         rowspan= " rowspan=" & (usrtsk + 1) & " "
      celdas = ofv.GenCelda("","vt" & rowspan,"","","","","",USRDESC,"si")
      celdas = CELDAS & ofv.GenCelda("","vt COLSPAN=2 ","","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
 
      filas = filas & FILAS1

      celdas = ofv.GenCelda("","AA COLSPAN=3 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
      end if


   end if


 ''   Celda = ofv.GenCelda("","abc colspan=3 style='border-bottom-style:ridge;border-bottom-width:2;border-bottom-color:#f7aa22;' ","center","100%","","","","&nbsp;","si")
 ''   FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=REQsectoresBUS.asp?VASP=ES  "
     sinput = sinput  & "  TITLE='Afectar&nbsp;sector'  ><b>" & sImagenI3 & "&nbsp;Afectar&nbsp;sector</b></a>" 

 ''    sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34) & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34) & "none" & chr(34) & ";' class=tt href=REQsectoresBUS.asp?VASP=ES"
 ''    sinput = sinput  & " ><b>&nbsp;afectar&nbsp;sector&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
     celdas = celdas & ofv.GenCelda("","vt","","10%","","","",sInput,"si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","100%","","0","","0","0","","",fila,"si")

    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
 '   FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa  ","","98%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"





 




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>