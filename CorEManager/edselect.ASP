<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  TipObjeto = request.querystring("TipObjeto")
  CodObjeto = request.querystring("CodObjeto")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select ARCHIVO,titulo From ObjMain where CodObjeto=" & CodObjeto & " and TipObjeto='"
  Strsql = Strsql & TipObjeto & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    RtfObjetivo=rs(0)
    TITObjetivo=rs(1)
  else
    TITObjetivo="OBJETIVO"
    RtfObjetivo="Obj" & CodObjeto & ".rtf"
    strsql = "insert into objMain (codobjeto,tipobjeto,archivo,titulo) values(" & codobjeto & ",'"
    strsql = strsql & TipObjeto & "','" &  RtfObjetivo & "','OBJETIVO')"
    ofv.crearconsultaEx StrSql,cn,1,parametros
  end if 
  call ofv.cerrarconsulta(rs)

  Strsql = "Select ARCHIVO,titulo From ObjAlcance where CodObjeto=" & CodObjeto & " and TipObjeto='"
  Strsql = Strsql & TipObjeto & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    RtfAlcance=rs(0)
    TITAlcance=rs(1)
  else
    TITAlcance="ALCANCE"
    RtfAlcance="Alc" & CodObjeto & ".rtf"
    strsql = "insert into objAlcance (codobjeto,tipobjeto,archivo,titulo) values(" & codobjeto & ",'"
    Strsql = Strsql & TipObjeto & "','" &  RtfAlcance & "','ALCANCE')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)
  end if 
  call ofv.cerrarconsulta(rs)

  Strsql = "Select ARCHIVO,titulo From ObjConsid where CodObjeto=" & CodObjeto & " and TipObjeto='"
  Strsql = Strsql & TipObjeto & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    RtfConsid=rs(0)
    TITConsid=rs(1)
  else
    TITConsid="CONSIDERACIONES GENERALES"
    RtfConsid="Con" & CodObjeto & ".rtf"
    strsql = "insert into objConsid (codobjeto,tipobjeto,archivo,titulo) values(" & codobjeto & ",'"
    Strsql = Strsql & TipObjeto & "','" &  RtfConsid & "','CONSIDERACIONES GENERALES')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)
  end if 
  call ofv.cerrarconsulta(rs)

  Xfer = server.mappath(request.servervariables("SCRIPT_NAME"))
  XferTarget1 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "carpetaObjetivos\"
  XferTarget2 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetaAlcance\"
  XferTarget3 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetaConsideracion\"

  abrir1 = XferTarget1 & RtfObjetivo
  abrir2 = XferTarget2 & RtfAlcance
  abrir3 = XferTarget3 & RtfConsid

'  if ofv.explorador = "MSIE" then
     sAccion = "edtransfer.asp"
'  else
'     sAccion = "edtransferNSC.asp"
'  end if
  sAccion1 = sAccion & "?Codobjeto=" & CodObjeto & "&Tipobjeto=" & Tipobjeto & "&opc=1&archivo=" & abrir1
  sAccion2 = sAccion & "?Codobjeto=" & CodObjeto & "&Tipobjeto=" & Tipobjeto & "&opc=2&archivo=" & abrir2
  sAccion3 = sAccion & "?Codobjeto=" & CodObjeto & "&Tipobjeto=" & Tipobjeto & "&opc=3&archivo=" & abrir3

'  sHTML = "<html><body bgcolor=#ccbbaa>"
  sHTML = ofv.FormBody

' Boton actualizar Titulos
  sAccion = "EdSelectTit.asp?Codobjeto=" & CodObjeto & "&Tipobjeto=" & Tipobjeto
  sInput = ofv.GenerarInput("","Titulos","submit","","20","bt","")
  Formulario = ofv.GenForm("FormX","",sAccion,"_parent","post","",sInput,"si")
  Celda = ofv.GenCelda("","cca STYLE={border-style:outset;border-width:1;} ","left","33%","","","",Formulario,"si")

' Boton de Objetivo
  sInput = ofv.GenerarInput("",TITObjetivo,"submit","","20","bt","")
  Formulario = ofv.GenForm("Form1","",sAccion1,"Transfer","post","",sInput,"si")
  Celda = Celda & ofv.GenCelda("","cc STYLE={border-style:outset;border-width:1;} ","","10%","","","",Formulario,"si")

' Boton de Alcance
  sInput = ofv.GenerarInput("",TITAlcance,"submit","","20","bt","")
  Formulario = ofv.GenForm("Form2","",sAccion2,"Transfer","post","",sInput,"si")
  Celda = Celda & ofv.GenCelda("","cc STYLE={border-style:outset;border-width:1;} ","","10%","","","",Formulario,"si")

' Boton de Consideraciones
  sInput = ofv.GenerarInput("",TITConsid,"submit","","20","bt","")
  Formulario = ofv.GenForm("Form3","",sAccion3,"Transfer","post","",sInput,"si")
  Celda = Celda & ofv.GenCelda("","cc STYLE={border-style:outset;border-width:1;} ","","10%","","","",Formulario,"si")

' Boton de Test
  sAccion = "CoConsideraciones.asp?Codobjeto=" & CodObjeto & "&Tipobjeto=" & Tipobjeto
  sAccion = sAccion & "&Menu=PORESTADO&Sector=3&origen=ED"
  sInput = ofv.GenerarInput("","Test","submit","","20","bt","")
  Formulario = ofv.GenForm("FormT","",sAccion,"_blank","post","",sInput,"si")
  Celda = Celda & ofv.GenCelda("","cca STYLE={border-style:outset;border-width:1;} ","right","33%","","","",Formulario,"si")

  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML & ofv.GenTabla("","aaa STYLE={border-style:solid;border-width:1;border-color:#000080;}","center","100%","","0","","-1","1","","",Fila,"si")

  sHTML = shtml & "</body></html>"
end if

response.write shtml

%>