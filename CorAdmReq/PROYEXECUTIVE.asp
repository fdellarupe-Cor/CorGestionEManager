<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  sHTML = ofv.formHeader(" ")


    set cn = ofv.conectar(ofv.strconn6)

   StrSql = "Select * From PROYECTO where codigo = '" & session("NCliente") & "' "
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if NOT rs.eof then
''      estado = RS(11) 
''      origin = RS(13) 
''      asigna = RS(14) 
''      DELEGA = RS(17) 

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

    sLink = sImagen & "&nbsp;RESUMEN&nbsp;EJECUTIVO&nbsp;DEL&nbsp;PROYECTO"
    Celda = ofv.GenCelda("","ttt colspan=2","","100%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sLink = "<hr width=100% size=1 color=#993333 >"
    Celda = ofv.GenCelda("","vt colspan=2","","100%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<B><FONT FACE=WINGDINGS COLOR=#BB3333 SIZE=1>l</FONT></B>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3 = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
    sImagen4so = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=2>&#0128;</FONT></B>"
    sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=2>&#0128;</FONT></B>"
    sImagen4sA = "<B><FONT FACE=WEBDINGS COLOR=#BBBBBB SIZE=2>&#0128;</FONT></B>"
    sImagen4n = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>x</FONT></B>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"
    sImagen6SS = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagen6N = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=3>3</FONT></B>"
    sImagen6OT = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=3>&#0254;</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B>"
    sImagenut = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=3>&#0192;</FONT></B>"
    sImagendu = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=3>6</FONT></B>"
    sImagenfi = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>R</FONT></B>"
    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0236;</FONT></B>"
    sImagenS = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenES = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=3>&#0232;</FONT></B>"
    sImagenP = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=3>2</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=3>&#0237;</FONT></B>"



''    sImagen = ofv.GenImage("","aa style='border-style:outset;border-color:#336699;' ","","","imagenes/botrep.jpg","","")

    sImagen = sImagen1N

    sHref = "PROYEXTXR.asp"
    sLink = ofv.GenLink("","vt",sHref,"Cont","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","CENTER","10%","","","",sLink,"si")
    sLink = ofv.GenLink("","cc",sHref,"Cont","","","TAREAS POR RESPONSABLE","si")
    Celda = Celda & ofv.GenCelda("","vt","","80%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

    sHref = "PROYEXTHR.asp"
    sLink = ofv.GenLink("","vt",sHref,"Cont","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","CENTER","10%","","","",sLink,"si")
    sLink = ofv.GenLink("","cc",sHref,"Cont","","","TOTAL HORAS POR RESPONSABLE","si")
    Celda = Celda & ofv.GenCelda("","vt","","80%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

    sHref = "PROYEXTHP.asp"
    sLink = ofv.GenLink("","vt",sHref,"Cont","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","CENTER","10%","","","",sLink,"si")
    sLink = ofv.GenLink("","cc",sHref,"Cont","","","TOTAL HORAS DEL PROYECTO","si")
    Celda = Celda & ofv.GenCelda("","vt","","80%","","","",sLink,"si")
  '  Fila = Fila & ofv.GenRow("","","","","30","","",Celda,"si")



  END IF

  sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",fila,"si")


  sHTML = sHTML & "</center></body></html>"
end if

  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)

Response.Write sHTML

%>