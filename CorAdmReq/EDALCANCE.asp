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

  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 


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



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;ALCANCE DEL REQUERIMIENTO","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if RS(9) = 0 then
      celdas = ofv.GenCelda("","ut","","10%","","","","TEMA/SUBTEMA:","si")
      sInput = "SIN DETERMINAR"
      celdas = celdas & ofv.GenCelda("","uti","","60%","","","",sInput,"si")




     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=EDALCANCETS.asp  "
     sinput = sinput  & "  TITLE='ASIGNAR'  ><b>" & sImagenI3 & "&nbsp;ASIGNAR</b></a>" 


      celdas = celdas & ofv.GenCelda("","ut  ","center","10%","","","",sinput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    else
      XTST = SPLIT(GETTEMASUBTEMA(RS(9)),CHR(10))
      SINPUT1 = "TEMA:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>" & XTST(0) & "</B>"
      SINPUT2 = "SUBTEMA:&nbsp;&nbsp;<B>" & XTST(1) & "</B>"
      

      celdas = ofv.GenCelda("","ut COLSPAN=2","","90%","","","",sInput1,"si")
   '   sinput = "<a class=cc style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34) & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34) & "none" & chr(34) & ";' class=tt href=EDALCANCETS.asp"
   '   sinput = sinput  & " ><b>MODIFICAR</b></a>" 

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=EDALCANCETS.asp  "
     sinput = sinput  & "  TITLE='MODIFICAR'  ><b>" & sImagenI3 & "&nbsp;MODIFICAR</b></a>" 


      celdas = celdas & ofv.GenCelda("","ut rowspan=2 ","center","10%","","","",sinput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","ut COLSPAN=2","","","","","",sInput2,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  ''  celdas = ofv.GenCelda("","AA COLSPAN=2","","80%","","","","&nbsp;","si")
      sinput = "<a class=cc style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34) & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34) & "none" & chr(34) & ";' class=tt href=EDALCANCETS.asp"
      sinput = sinput  & " ><b>MODIFICAR</b></a>" 
 ''     celdas = celdas & ofv.GenCelda("","TT  ","center","","","","",sinput,"si")
      celdas = celdas & ofv.GenCelda("","ut rowspan=2 ","center","","","","",sinput,"si")



 '     filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    end if

      celdas = ofv.GenCelda("","vt colspan=3 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","center","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    'Encabezado DESCRIPCION


    Celda = ofv.GenCelda("","ut colspan=3","","100%","","","","DESCRIPCION DEL REQUERIMIENTO","si")
    FilaS =  filaS & ofv.GenRow("","","","","20","","",Celda,"si")

      nCol = 80

  if len(wdescri) = 0 then wdescri = "Sin Descripcion" 

  if not  isnull(rs(10)) then 
     wdescri = rs(10)
  end if 




    t3b = "<font class=at><textarea class=dt  style='background-color:transparent;padding-top:2px;' "
    t3b = t3b & "name=" & rs(10).name & " cols=95 rows=20 "
    if bModificar then
      t3b = t3b & " onblur=Usuarios.submit() "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 1950) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    else
      t3b = t3b & " readonly >"
    end if
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","",t3b,"si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=REQUERIMIENTO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if







    filas = ofv.GenRow("","","","","","","",formulario0,"si")
    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"

 
  IF session("ModMain2") = "CL" THEN
     sinput = "<a  class=tt href=Clientes.asp"
     sinput = sinput  & " ><b>&nbsp;Volver&nbsp;a&nbsp;Cliente&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

  ELSE
     sinput = "&nbsp;"
     celdas =  ofv.GenCelda("","","","1","","","",sInput,"si")
  END IF


  celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
  filas = ofv.GenRow("","","","","","","",celdas,"si")
  sHTML = sHTML  & ofv.GenTabla("","aa ","","98%","","0","","0","0","","",filas,"si")

  end if



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


''''''''''''''''''''''

  sHTML   = sHTML & CorPnlCOT()

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clpanel();' ", "S", "A", "N", "indFrm", "")

  Agregar = Celt1
  Buscar  = ""
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

''''''''''''''''''''''


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GETTEMASUBTEMA(SSUBTEMA)

  xTEMA = 0
  xSUBTEMA = "SIN ASIGNAR"

  SQL2 = "SELECT * "
  SQL2 = SQL2 & " FROM csbuscasubtema where "
  SQL2 = SQL2 & " codPRODSERV = " & SSUBTEMA
  set rsX = ofv.crearconsultaEx(SQL2,cn3,1,parametros)
  if NOT rsX.eof then
     xTEMA = RSX(3)
     xSUBTEMA = RSX(1)
     SQL1 = "SELECT * From Tema WHERE CODTEMA = " & xTEMA
     set rsX = ofv.crearconsultaEx(SQL1,cn3,1,parametros)
     if NOT rsX.eof then
        xTEMA = RSX(1)
     END IF

  END IF

  GETTEMASUBTEMA = XTEMA & CHR(10) & XSUBTEMA

END FUNCTION


%>