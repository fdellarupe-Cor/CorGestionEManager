<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.MENUHeader(""," ONLOAD=" & chr(34) & "document.title='Cor OnTime - Vista de Planificaciones';" & chr(34) & " ")



  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar




 VPLAN = request.querystring("PLAN")






''    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#FFFFFF SIZE=2>l</FONT>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen0 = "<FONT FACE=WINGDINGS COLOR=#6699cc  SIZE=2>2</FONT>"
    sImagen3 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0166;</FONT></B>"
    sImagen4n = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>N</FONT></B>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"
    sImagen6SS = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=3>&#0232;</FONT></B>"
    sImagen6N = "<FONT COLOR=#993333 FACE=WEBDINGS SIZE=2>4</FONT>"
    sImagen6M = "<FONT COLOR=#993333 FACE=WEBDINGS SIZE=2>8</FONT>"
    sImagen6OT = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>r</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=2>r</FONT></B>"
    sImagenSTN = "<B><FONT FACE=WINGDINGS COLOR=#336699 SIZE=3>6</FONT></B>"
    sImagenSTC = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0184;</FONT></B>"
    sImagenE0 = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=4>6</FONT></B>"
    sImagenE1 = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"
    sImagenE2 = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=4>&#0251;</FONT></B>"
    sImagenCC = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=4>2</FONT></B>"
 




    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

      celdas = ofv.GenCelda("","cc ","","","","","",sImagen & "&nbsp;<BIG><B>PLANIFICACIONES DEL PROYECTO</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      FilaS1 = ofv.GenTabla("","","","100%","","0","","-2","2","","",filas,"si")


      filas = ""



      SINPUT = "&nbsp;"
      celdas = ofv.GenCelda("","vt colspan=3 STYLE='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:#f7aa22;' ","CENTER","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  StrSql = "Select * From PROYPLAN where ID = " & VPLAN & " "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs1.eof then

      SINPUT = "--------------"
      celdas = ofv.GenCelda("","vt colspan=3 ","","","","","",SINPUT,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  ELSE



      SINPUT = "PLAN "
      celdas = ofv.GenCelda("","vt ","","20%","","","",SINPUT,"si")

      SINPUT = ":"
      celdas = celdas & ofv.GenCelda("","VT ","","4%","","","",SINPUT,"si")

      SINPUT = RS1(2) & ".  " & RS1(3)
      celdas = celdas & ofv.GenCelda("","VT ","","70%","","","",SINPUT,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      SINPUT = "CREACION"
      celdas = ofv.GenCelda("","vt ","","","","","",SINPUT,"si")

      SINPUT = ":"
      celdas = celdas & ofv.GenCelda("","VT ","","","","","",SINPUT,"si")

      SINPUT = cdate(RS1(4) + clng(cdate("01-01-1900")))
      celdas = celdas & ofv.GenCelda("","VT ","","","","","",SINPUT,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



      SINPUT = "RESPONSABLE"
      celdas = ofv.GenCelda("","vt ","","","","","",SINPUT,"si")

      SINPUT = ":"
      celdas = celdas & ofv.GenCelda("","VT ","","","","","",SINPUT,"si")

    if RS1(6) = "NO" then
      sInput = "SIN RESPONSABLE"
      celdas = celdas & ofv.GenCelda("","vT","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & RS1(6) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sImagenO & " <B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","VT","","","","","",SINPUT,"si")
    END IF

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      SINPUT = "ESTADO"
      celdas = ofv.GenCelda("","vt ","","","","","",SINPUT,"si")

      SINPUT = ":"
      celdas = celdas & ofv.GenCelda("","VT ","","","","","",SINPUT,"si")

      IF RS1(7) = 1 THEN
         SINPUTE = sImagen6N & "&nbsp;VIGENTE"
      ELSEIF RS1(7) = 0 THEN
         SINPUTE = sImagen6N & "&nbsp;PENDIENTE"
      ELSEIF RS1(7) = 2 THEN
         SINPUTE = sImagen6M & "&nbsp;REEMPLAZADO"
      ELSE
         SINPUTE = sImagen6T & "&nbsp;CANCELADO"
      END IF


      SINPUT = SINPUTE
      celdas = celdas & ofv.GenCelda("","VT ","","","","","",SINPUT,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



      celdas = ofv.GenCelda("","vt COLSPAN=3 style='border-BOTTOM-style:ridge;border-BOTTOM-width:2;border-BOTTOM-color:#f7aa22;'","CENTER","","","","","<BR>OBSERVACIONES","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      nCol = 80

      wdescri = RS1(8) & "-"


    t3b = "<font class=at><textarea class=dt  style='background-color:transparent;padding-top:2px;' "
    t3b = t3b & "name=VPOBSERV cols=" & nCol & " rows=6 "
    t3b = t3b & "  >"
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    sInput = t3b

      celdas = ofv.GenCelda("","vt COLSPAN=3","CENTER","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





  END IF 

  call ofv.cerrarconsulta(rs1)



    sHTML = sHTML & FILAS1 
    sHTML = sHTML & ofv.GenTabla("","aaA  style='filter:shadow(color=#cccccc,direction=125);' ","","98%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"


  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML





%>