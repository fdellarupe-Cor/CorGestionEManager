<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

nAccion     = request.querystring("Accion")
nProced     = request.querystring("Codproced")
nEsquema    = request.querystring("Esquema")
nControl    = request.querystring("Control")
nCodigo     = request.querystring("Codobjeto")
sEven       = request.querystring("Codeven")
sEven       = ofv.convertircar(sEven,"_"," ")
sTipo       = request.querystring("Tipobjeto")
Tabla       = request.querystring("Tabla")
sResp       = request.querystring("tiposeg")
Calificador = request.querystring("Calificador")
Volver      = request.querystring("Volver")
Volver      = ofv.convertircar(Volver,"_","&")
sUsuario    = session("Usuario")
'segckt		=	request.querystring("segckt")
'response.write request.querystring
on error resume next

set cn = ofv.conectar(ofv.strconn2)

wkspsep = ","

strsql = "Select * From " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

if Tabla <> "Objaccioncontrol" and Tabla <> "ORRSKCIRCUITOSDKDET" and ucase(Tabla) <> "ORUSRLST" and lcase(tabla) <>"novegenpardist" then
  strsql = "Insert into " & tabla & " (Codobjeto,Tipobjeto"
end if

if Tabla = "ORRSKCIRCUITOSDKDET"  then
  strsql = "Insert into " & tabla & " ("
  wkspsep = ""
end if

if  ucase(TABLA) = "ORUSRLST" then
  strsql = "Insert into " & tabla & " ("
  wkspsep = ""
end if
if  ucase(TABLA) = "NOVEGENPARDIST" then
  strsql = "Insert into " & tabla & " ("
  wkspsep = ""
end if


if Tabla = "Objaccion" then
  strsql = strsql & ",Codproced,modo"
  vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form("tipoaccion"),"_"," ")))
  if vNuevovalor <> "" then
    sModo = "XX"
    sql = "Select modo From Tipoaccalt Where tipoaccionalt = " & vNuevoValor
    set rsModo = ofv.crearconsultaEx(sql,cn,1,volver)
    if not rsModo.eof then sModo = rsModo(0)
    call ofv.cerrarconsulta(rsModo)
  end if
elseif Tabla = "Objesquema2" then
  strsql = strsql & ",codesquema"
elseif Tabla = "Objcontrol2" then
  strsql = strsql & ",control"
elseif tabla = "Objaccioncontrol" then
  sql = "SELECT max(secuencia) FROM Objaccioncontrol "
  sql = sql & "WHERE Objaccioncontrol.idaccion= '" & nAccion & "'"
  set rsSecuencia = ofv.crearconsultaEx(Sql,cn,1,volver)
  nSecuencia = 1
  if not rsSecuencia.eof then
    if not isnull(rsSecuencia(0)) then nSecuencia = rsSecuencia(0) + 1
  end if
  call ofv.cerrarconsulta(rsSecuencia)
  strsql = "Insert into " & tabla & " (Codproced,secuencia,idaccion"
end if

wkspsep2 = wkspsep

For Each x In Request.form
  strsql = strsql & wkspsep2 & x
  wkspsep2 = ","
Next

if Tabla = "OBJwflDET" then
'strsql=strsql & wkspsep2 & "codEmp"
end if
if Tabla = "OBJSDKdet" then
'strsql=strsql & wkspsep2 & "codEmp"
end if

wkspsep2 = wkspsep


if Calificador <> "" then strsql = strsql & ",Calificador"
if Tabla <> "Objaccioncontrol" and Tabla <> "ORRSKCIRCUITOSDKDET" AND ucase(TABLA) <> "ORUSRLST" and ucase(tabla) <>"NOVEGENPARDIST" then
  strsql = strsql & ") Values ('" & nCodigo & "','" & sTipo & "'"
elseif Tabla = "ORRSKCIRCUITOSDKDET" then
  strsql = strsql & ") Values ("
elseif ucase(tabla)="ORUSRLST" then
  strsql = strsql & ") Values ("
elseif ucase(tabla)="NOVEGENPARDIST" then
  strsql = strsql & ") Values ("
else
  strsql = strsql & ") Values ('" & nProced & "','" & nSecuencia & "','" & nAccion & "'"
end if
if Tabla = "Objaccion" then
  strsql = strsql & ",'" & nProced & "','" & sModo & "'"
elseif Tabla = "Objesquema2" then
  strsql = strsql & "," & nEsquema
elseif Tabla = "Objcontrol2" then
  strsql = strsql & "," & nControl
end if

ValorNulo = False
TipoValido = True
For Each YY In Request.form
X = LCASE(YY)


  if x = "dettexto" or (tabla = "Objcontrol2" and x = "descri") then
     vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
  elseif UCASE(tabla)="OBJSDKDET" OR UCASE(tabla)="ORRSKCIRCUITOSDKDET" THEN
     vNuevoValor =UCASE(CStr(Request.Form(x)))
  ELSEif UCASE(tabla)="OBJDIST"  THEN
     vNuevoValor =UCASE(CStr(Request.Form(x)))
  ELSEif ((ucase(tabla)= "OBJTABDEF" OR ucase(tabla)= "OBJTABBLK" OR ucase(tabla)= "OBJTABCON")  and session("APLMDCLI")="BANELCO") THEN
   vNuevoValor =CStr(Request.Form(x))
  ELSEif UCASE(tabla)="OBJWFLDET" AND X="usr" THEN
      vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
  else

  '' 07112014 - anulada por mayuscula minuscula
  '' vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
  '' 07112014 - anulada por mayuscula minuscula

     vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))


  end if

  vNuevoValor = replace(vNuevoValor,"'","''")
  s = ofv.QueEsx(rs(x).type)
  if rs(x).type = 200 then s = "'"
  if rs(x).type > 199 and rs(x).type < 300 then
     s = "'"
  elseif rs(x).type = 135 then
     s = ""
  end if
  if Tabla = "Objaccion" then
    ValorNulo = ofv.EsNulo(rs(x),vNuevoValor)
    if ValorNulo or vNuevoValor = "NSNC" then
      if rs(x).name = "tipoaccion" then
        TipoValido = False
        exit for
      end if
      if s = "'" then
        strsql = strsql & wkspsep2 & "''"
      else
        strsql = strsql & wkspsep2 & "0"
      end if
    else
      strsql=  strsql & wkspsep2 & s & vNuevoValor & s
    end if
    ValorNulo = False
  else
    if x <> "codSector" then
      if ucase(vNuevoValor) = "TRUE" then
        vNuevoValor = 1
      elseif ucase(vNuevoValor) = "FALSE" then
        vNuevoValor = 0
      end if
      ValorNulo = ofv.EsNulo(rs(x),vNuevoValor)
      if ValorNulo then exit for
      TipoValido = ofv.ValidarTipo(rs(x),vNuevoValor)

'      if not TipoValido then exit for

    end if
    if vNuevoValor <> "NSNC" then
      strsql=  strsql & wkspsep2 & s & vNuevoValor & s
    else
      if s = "'" then
        strsql = strsql & wkspsep2 & "''"
      else
        strsql = strsql & wkspsep2 & "0"
      end if
    end if
  end if

  wkspsep2 = ","

Next



if Calificador <> "" then strsql = strsql & "," & Calificador

if volver <> "" then
  if tabla <> "Objaccioncontrol" then
    if right(volver,3) = "asp" then
      volver = volver & "?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    else
      volver = volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if
  end if
end if

 if tabla = "OBJwflDET" then
 ' 	strsql = strsql & wkspsep2 & session("SETIDEMPRE")
  end if
 if tabla = "OBJSDKdet" then
'   	strsql = strsql & wkspsep2 & session("SETIDEMPRE")
  end if


if not TipoValido or ValorNulo then
  on error resume next
'  response.write strsql
  response.write ofv.ErrorTipo(volver,Request.Form(x),rs(nx).name)
  if err.number <> 0 then response.write ofv.ErrorTipo(volver,"","")
else
  strsql = strsql & ")"

'  response.write strsql
if sTipo = "SN" and ucase(tabla) = "PEDIDOSDKDET" then
	strsql = "update " & tabla & " set codemp=" & request.form("codemp") & ",grpusr='" & SESSION("FBUSR") & "' where tipobjeto='" & sTipo & "' and tiposeg='" & sResp & "'"
end if

	call ofv.crearconsultaEx(StrSql,cn,1,volver)

  if cn.errors.count = 0 then
     if Ucase(tabla) <> "OBJWFLDET"  and Ucase(tabla) <> "ORUSRLST" then
        strsql = "Update Objprty Set aprobo = '" & sUsuario & "' Where codobjeto = " & nCodigo
        call ofv.crearconsultaEx(StrSql,cn,1,volver)
     end if

     if ucase(Tabla) = "OBJACCION" or ucase(Tabla) = "OBJACCIONCONTROL" or ucase(Tabla) = "OBJESTR" or ucase(Tabla) = "OBJESTDET" or ucase(Tabla) = "OBJESQUEMA" or ucase(Tabla) = "OBJESQUEMA2" or ucase(Tabla) = "OBJCONTROL" or ucase(Tabla) = "OBJCONTROL2" or ucase(Tabla) = "OBJEVEN" or ucase(Tabla) = "OBJSECTOR" or ucase(Tabla) = "OBJFAQ" or ucase(Tabla) = "OBJHTM" or ucase(Tabla) = "OBJPROCED"  or ucase(Tabla) = "OBJTABDEF" then

        vFECHAHOY = clng(DATE) - clng(cdate("01-01-1900"))
        xxxx      = GETFECH()
        vHORAHOY  = XLHORDAT

        strsqlx = "Update " & Tabla & " Set modi = 1, tipmodi = 'A', MODAUTOR = '" & SESSION("USUARIO") & "',"
        strsqlx = strsqlx & " MODFECHA = '" & vFECHAHOY & "', MODHORA = '" & vHORAHOY & "'"
        strsqlx = strsqlx & " Where objid = " & ID
        call ofv.crearconsultaEx(StrSqlx,cn,1,volver)

     end if


 'response.write "<br><br>" & volver & "<br><br>"
 response.redirect volver
  else
    response.write ofv.ErrorDB(cn,volver)

  end if
end if
response.write strsql
call ofv.cerrarconsulta(rs)
call ofvcerrarconn(cn)

%>