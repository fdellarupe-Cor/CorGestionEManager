<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="../IncFile/CorWKFMANAGER.inc" -->
<%


response.expires=0

Vuelta    = request.querystring("Vuelta")
nAccion   = request.querystring("Accion")
nProced   = request.querystring("Codproced")
nCodigo   = request.querystring("Codobjeto")
nConti    = request.querystring("Conti")
sTipo     = request.querystring("Tipobjeto")
sEliminar = request.querystring("Eliminar")
sopcwfl   = request.querystring("opc")
ID        = request.querystring("ID")
ID        = ofv.convertircar(ID,"*"," ")
Tabla     = request.querystring("Tabla")
sRef      = request.querystring("Referencia")
Volver    = request.querystring("Volver")
Volver    = ofv.convertircar(Volver,"_","&")
sUsuario  = session("Usuario")

  '  response.write "Volver request: " & volver & "<br>"

'''on error resume next
'response.write request.form
set cn = ofv.conectar(ofv.strconn2)


if sEliminar = "SI" then
  strsql = "Update " & tabla & " Set modi = 1, tipmodi = 'E' Where objid = " & ID
  call ofv.crearconsultaEx(StrSql,cn,1,volver)
  if cn.errors.count = 0 then
    if tabla = "Objaccioncontrol" then
      Volver = Volver & "?Codproced=" & nProced & "&Codobjeto=" & nCodigo & "&Accion="
      Volver = Volver & nAccion & "&Tipobjeto=" & sTipo & "&volver=" & vuelta
    elseif Ucase(tabla) = Ucase("Objaccion") then
           strsql = "select * from " & tabla & " Where objid = " & ID
           set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
           if not rs.eof then
              if clng(rs(19)) > 0 then
                 strsql = "delete objlnk Where objid = " & rs(19)
                 call ofv.crearconsultaEx(StrSql,cn,1,volver)

                 strsql = "Update " & tabla & " Set rellink = 0 Where objid = " & ID
                 call ofv.crearconsultaEx(StrSql,cn,1,volver)
              end if
           end if
           call ofv.cerrarconsulta(rs)
           volver = volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    else
      volver = volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if
    response.redirect volver
  else

    response.write ofv.ErrorDB(cn,volver)
  end if
else

' response.write "id: " & ID

strsql = "select * from " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

if Ucase(tabla) = "OBJPERFILDOC" then
   strsql = "select * from " & tabla & " where objid =" & id
else
   s = ofv.QueEsx(rs(0).type)
   if rs(0).type = 200 then s = "'"
   strsql = "select * from " & tabla & " where " & rs(0).name & "=" & s & id & s
end if
'response.write strsql
call ofv.cerrarconsulta(rs)
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


strsql = "update " & tabla & " set "

if Tabla = "Objaccion" then
  if sRef <> "SI" then
    vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form("tipoaccion"),"_"," ")))
    if vNuevoValor = "NSNC" then vNuevoValor = 0
    sql = "Select modo From Tipoaccalt Where tipoaccionalt = " & vNuevoValor
    set rsModo = ofv.crearconsultaEx(sql,cn,1,volver)
    m = ""
    if not rsModo.eof then m = rsModo(0)
    call ofv.cerrarconsulta(rsModo)
    strsql =  strsql & "modo = '" & m & "', "
    vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form("codcomp"),"_"," ")))
    if vNuevoValor = "NSNC" then vNuevoValor = 0
    sql = "Select tipocomp From Complementos Where codcomp = " & vNuevoValor
    set rsTipocomp = ofv.crearconsultaEx(sql,cn,1,volver)
    tc = ""
    if not rsTipocomp.eof then tc = rsTipocomp(0)
    call ofv.cerrarconsulta(rsTipocomp)
    strsql =  strsql & "tipocomp = '" & tc & "', "
  end if
'elseif ucase(Tabla)="OBJEVEN" THEN
'   strsql=strsql & " CODPROCESO=ISNULL(CODPROCESO,'') , "
 ' ASAS

end if

SQL="SELECT VERSION FROM OBJPRTY WHERE CODOBJETO=0" & nCodigo

set rsVERSION = ofv.crearconsultaEx(sql,cn,1,volver)

XSVERSION="N"

IF NOT rsVERSION.EOF THEN

IF rsVERSION(0)="O" THEN
  XSVERSION="N"
ELSE
  XSVERSION="S"
END IF
END IF
CALL OFV.CERRARCONSULTA(rsVERSION)

IF XSVERSION="S" THEN
if ucase(Tabla) = "OBJACCION" or ucase(Tabla) = "OBJACCIONCONTROL" or ucase(Tabla) = "OBJESTR" or ucase(Tabla) = "OBJESTDET" or ucase(Tabla) = "OBJESQUEMA" or ucase(Tabla) = "OBJESQUEMA2" or ucase(Tabla) = "OBJCONTROL" or ucase(Tabla) = "OBJCONTROL2" or ucase(Tabla) = "OBJEVEN" or ucase(Tabla) = "OBJSECTOR" or ucase(Tabla) = "OBJFAQ" or ucase(Tabla) = "OBJHTM" or ucase(Tabla) = "OBJPROCED"  or ucase(Tabla) = "OBJTABDEF"   or ucase(Tabla) = "OBJTABBLK"  or ucase(Tabla) = "OBJTABCON" then

   vFECHAHOY = clng(DATE) - clng(cdate("01-01-1900"))

   xxxx = GETFECH()
   vHORAHOY  = XLHORDAT

   strsqlx = "Update " & Tabla & " Set modi = 1, tipmodi = 'M', MODAUTOR = '" & SESSION("USUARIO") & "',"
   strsqlx = strsqlx & " MODFECHA = '" & vFECHAHOY & "', MODHORA = '" & vHORAHOY & "'"
   strsqlx = strsqlx & " Where objid = " & ID
   call ofv.crearconsultaEx(StrSqlx,cn,1,volver)

end if
END IF


vinofv    = "N"
vinover   = "N"
bEjecutar = False
sEstado1  = request.form("estado")

For Each yy In Request.form

x = lcase(yy)

'response.write "x: " & x & " " & Request.Form(x) & "|<BR> "

  if x = "modi" then
     if Ucase(Request.Form(x)) = "FALSE" or Request.Form(x) = "0" then
        strsqlm = "Update " & tabla & " Set modi = 0, tipmodi = ' ' Where objid = " & ID
        call ofv.crearconsultaEx(StrSqlm,cn,1,volver)
     end if
  else

''ObjFAQ

  if not (sEstado1 =  5 and  x = "fecvigen") then
     if (Ucase(tabla) = "OBJDSCABR" and x = "descri") or (Ucase(tabla) = "OBJDSCABR" and x = "titulo") then
        vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
     elseif (Ucase(tabla) = "OBJACCION" and (x = "referencia" or x = "referencia2" or x = "referencia3" or x = "reflink" or x = "descriacc")) then
     '   vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
        vNuevoValor = CStr(Request.Form(x))
     elseif (Ucase(tabla) = "OBJPROCED" and (x = "referencia" )) then
        vNuevoValor = CStr(Request.Form(x))
     elseif (Ucase(tabla) = "OBJACCIONCONTROL" and (x = "referencia")) then
        vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
     elseif (Ucase(tabla) = "OBJETOS" and (x = "descriabr")) then
        vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
     elseif (Ucase(tabla) = "OBJESQUEMA" and (x = "refesq")) then
        vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
     elseif (Ucase(tabla) = "OBJESQUEMA2" and (x = "refesq2")) then
        vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
     elseif (Ucase(tabla) = "OBJCONTROL" and (x = "referencia")) then
        vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
     elseif (Ucase(tabla) = "OBJCONTROL2" and x = "descri") then
        vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
     elseif (Ucase(tabla) = "OBJESTDET" and (x = "dettexto" or x = "dettitulo")) then
        vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
     elseif (Ucase(tabla) = "OBJFAQ" and (UCASE(CSTR(x)) = "REDURL" )) then
        vNuevoValor = CStr(Request.Form(x))
     elseif (Ucase(tabla) = "OBJPRTY" and (UCASE(CSTR(x)) = "PRTYLINK"  OR UCASE(CSTR(x)) = "PRTYLINK2" )) then
        vNuevoValor = CStr(Request.Form(x))
     elseif (Ucase(tabla) = "OBJPRTYADV" and (UCASE(CSTR(x)) = "PRTYLINK"  OR UCASE(CSTR(x)) = "PRTYLINK2" )) then
        vNuevoValor = CStr(Request.Form(x))
     elseif x = "detimagen" or ( x="archivo" and Ucase(tabla) = "OBJHTM" ) then
        vNuevoValor = CStr(Request.Form(x))
     elseif (Ucase(tabla) = ucase("Objestdet") and x = "dettexto") then
        vNuevoValor = (CStr(Request.Form(x)))
      elseif (Ucase(tabla) = ucase("OBJdistPrty") and x = "texto") then
        vNuevoValor =(CStr(Request.Form(x)))

     elseif ((ucase(tabla)= "OBJTABDEF" OR ucase(tabla)= "OBJTABBLK" OR ucase(tabla)= "OBJTABCON")  and session("APLMDCLI")="BANELCO") then
       vNuevoValor= CStr(ofv.convertircar(Request.Form(x),"_"," "))
     else
        vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
     end if
     vNuevoValor = replace(vNuevoValor,"'","''")

'response.write "x: " & x & " " & Request.Form(x) & "|2<BR> " &tabla

if tabla = "Objproced" and x="codperfil" then
 if len(Request.Form(x)) =0 then vNuevoValor=NULL
end if

  if tabla = "Objtabcon" then
     if x = "tabconsec" then
        if not (isnumeric(vNuevoValor)) then vNuevoValor = 0
     elseif  x <> "tabcontipo" then
        if len(vNuevoValor) = 0 then vNuevoValor = " "
     end if
  end if

  if vNuevoValor = "FALSE"     then vNuevoValor = "0"
  if vNuevoValor = "TRUE"      then vNuevoValor = "-1"
  if vNuevoValor = "FALS0"     then vNuevoValor = "0"
  if vNuevoValor = "VERDADERO" then vNuevoValor = "-1"

  if (x = "fecemis" or x = "fecbaja" or x = "fecvigen" or x = "fecvto" or x = "fecacta") then

     if vNuevoValor = "" then
        vNuevoValor = clng(cdate("01-01-1900")) - nFecini
    else
      ''''''  vNuevoValor = clng(cdate(vNuevoValor)) - nFecini

        vNuevoValor = clng(cdate(vNuevoValor))

     end if
  end if

  if  x = "version"  and Ucase(tabla) = "OBJPRTY" and Ucase(Request.Form(x)) = "V" then
      vinover = "S"
  end if

  if  x = "nrover"  and Ucase(tabla) = "OBJPRTY"  then
      vinover = "N"
  end if

 ''    response.write x & " 1 " & vNuevoValor & (vNuevoValor = "") & "<br>"

  if  x = "fecvigen"  then
     vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
      vinofv = "S"

     if vNuevoValor = "" then
        vNuevoValor = clng(cdate("01-01-1900")) - nFecini
    else
      '''''  vNuevoValor = clng(cdate(vNuevoValor)) - nFecini
        vNuevoValor = clng(cdate(vNuevoValor))
     end if
  end if

  if x = "codcuenta" then
    i = instr(vNuevoValor,",")
    if i > 0 then
      sCuenta = left(vNuevoValor,i-1)
      sql = "Select objid From Objcuenta Where descri = '" & sCuenta & "'"
      set rsCuenta = ofv.crearconsultaEx(Sql,cn,1,volver)
      if not rsCuenta.eof then vCuenta = rsCuenta(0)
      call ofv.cerrarconsulta(rsCuenta)
      if vCuenta = rs(x).Value then
        nCuenta = trim(mid(vNuevoValor,i+1))
        sql = "Select objid From Objcuenta Where numero = '" & nCuenta & "'"
        set rsCuenta = ofv.crearconsultaEx(Sql,cn,1,volver)
        if not rsCuenta.eof then vCuenta = rsCuenta(0)
        call ofv.cerrarconsulta(rsCuenta)
      end if
      vNuevoValor = vCuenta
    end if
  elseif x = "codevento" and Tabla = "Objesquema" then
    sql = "Select objid From Objeven Where descri = '" & vNuevoValor & "'"
    set rsEvento = ofv.crearconsultaEx(Sql,cn,1,volver)
    if not rsEvento.eof then vNuevoValor = rsEvento(0)
    call ofv.cerrarconsulta(rsEvento)
  end if

'response.write "x: " & x & " " & Request.Form(x) & " vNuevoValor:"  & vNuevoValor  & "|3<BR> "

 ' response.write x & "<br>"
 'and ((x <> "objurldsc" and x <> "objtitulo" and  x <> "objcopete"  and  x <> "objmensaje") and ucase(tabla)="OBJINFO")

 if (x = "prtylink" or x = "prtylink2") and tabla="Objprtyadv" then TipoValido=true

' response.write len(request.form(x)) & "      " & x & "<br>"

' TipoValido=true
 if x <> "prtylink" and x <> "prtylink2" and x <> "objdirectorio" and x <> "objurldsc" and x <> "objtitulo" and  x <> "objcopete"  and  x <> "objmensaje" then
  TipoValido = ofv.ValidarTipo(rs(x).name,vNuevoValor)

  if TipoValido = False then exit for
  if x <> "prtylink" and x <> "prtylink2" and x <> "instid" and x <> "referencia" and x <> "referencia2" and x <> "referencia3" and x <> "objdirectorio" and x <> "objurldsc" and x <> "objtitulo" and  x <> "objcopete"  and  x <> "objmensaje" then
    if lcase(trim(x)) = "perfilcont" then
      if clng(rs("sectorcont")) <> 0 then
        ValorNulo = ofv.EsNulo(rs(x),vNuevoValor)
        if ValorNulo = True then exit for
      end if
    else
      ValorNulo = ofv.EsNulo(rs(x),vNuevoValor)
      if ValorNulo = True then exit for
    end if
  end if
end if

 ' response.write x & " 3 " & tabla  & "<br>"

  s = ofv.QueEsx(rs(x).type)


  if rs(x).type = 200 then s = "'"

  if rs(x).type > 199 and rs(x).type < 300 then
     s = "'"
  elseif rs(x).type = 135 then
     s = ""
  end if
 'response.write "x: " & x & " x.type: " &  rs(x).type & " s: " & s & " <br>"
' response.write "strsql: " & strsql & " medio"
  if vNuevoValor <> "NSNC" then
  '  response.write vNuevoValor & " " & rs(x).Value & " <br>"
    if cstr(vNuevoValor & "") <> cstr(rs(x).Value & "") then
      bEjecutar = True
      strsql=  strsql & x & " = " & s & vNuevoValor & s & " ,"
    elseif isnull(rs(x).value) and vNuevoValor <> "" then
      bEjecutar = True
      strsql=  strsql & x & " = " & s & vNuevoValor & s & " ,"
    end if
  else
    bEjecutar = True
    if s = "'" then
      strsql =  strsql & x & " = '',"
    else
      if Tabla = "Objaccioncontrol" then
        strsql=  strsql & x & " = 999999,"
      else
        if rs(x).name = "tipoaccion" then
          TipoValido = False
          exit for
        end if
        strsql=  strsql & x & " = 0,"
      end if
    end if
  end if

'response.write "<BR>strsql: " & strsql & " 31<BR>"

end if
end if

'response.write "x: " & x & " " & Request.Form(x) & " vNuevoValor:"  & vNuevoValor  & " bEjecutar:" & bEjecutar & " TipoValido:" & TipoValido & " ValorNulo:" & ValorNulo & "|4<BR> "

Next

'response.write "<BR>strsql: " & strsql & " <BR>"


sEstado1 = request.form("estado")
vFECHAHOY = clng(DATE) - clng(cdate("01-01-1900"))


if vinofv <> "S" then
  if Tabla = "Objprty" and sEstado1 = 3 then strsql = strsql & "fecvigen = '" & vFECHAHOY & "',"
end if

if vinover = "S" then
   strsql = strsql & "nrover = '1',"
end if



if volver <> "" then
  if tabla = "Objaccioncontrol" then
    Volver = Volver & "?Codproced=" & nProced & "&Codobjeto=" & nCodigo & "&Accion="
    Volver = Volver & nAccion & "&Tipobjeto=" & sTipo & "&volver=" & vuelta
  else
  '  response.write "Volver Original: " & volver & "<br>"
    if right(volver,3) = "asp" then
      volver = volver & "?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    else
      SXvolver = instr(volver,"&Codobjeto")
      if SXvolver > 0 then
         volver = volver & " "
      else
         volver = volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      end if
    end if
  end if
end if



'response.write strsql
'''if SSSS = True then
'response.write TipoValido  & " paso2 " &   ValorNulo
if (TipoValido = False) or (ValorNulo = True) then

'response.write (TipoValido = False)  & "qqqq	   "  & (ValorNulo = True)  &  "   "& " paso11112"

  on error resume next
  response.write ofv.ErrorTipo(volver,Request.Form(x),rs(x).name)
  if err.number <> 0 then response.write ofv.ErrorTipo(volver,"","")
else

  s = ofv.QueEsx(rs(0).type)
  if rs(0).type = 200 then s = "'"


  if nConti = "SI" and Ucase(tabla) = "OBJACCIONCONTROL" then
     strsql = mid(strsql,1,len(strsql)-1) & " where idaccion =" & ofv.QueEsx(rs("idaccion").type) & nAccion & ofv.QueEsx(rs("idaccion").type)
  elseif Ucase(tabla) = "OBJPERFILDOC" then
     strsql = mid(strsql,1,len(strsql)-1) & " where objid = " & id
  else
     strsql = mid(strsql,1,len(strsql)-1) & " where " & rs(0).name & "=" & s & id & s
  end if
' response.write "strsql: " & strsql & " medio"

  if Tabla = "Objprty" then
    dFechaNula = "1/1/1900"

    sql = "Select codobjeto, tipobjeto From Objprty Where Objprty.objid = " & id
    set rsCodigo = ofv.crearconsultaEx(Sql,cn,1,volver)
    if not rsCodigo.eof then
           nCodigo = rsCodigo(0)
           ntipobj = rsCodigo(1)
    end if
    call ofv.cerrarconsulta(rsCodigo)



    sql = "SELECT Estados.estado FROM OBJETOS  OBJETOS_1, OBJETOS, "
    sql = sql & " RELAOBJ, OBJPRTY, Estados  "
    sql = sql & " where "
    sql = sql & " OBJETOS.codobjeto = RELAOBJ.codobjmaster and "
    sql = sql & " OBJETOS.tipobjeto = RELAOBJ.tipobjmaster and "
    sql = sql & " OBJETOS_1.codobjeto = RELAOBJ.codojasoc and "
    sql = sql & " OBJETOS_1.tipobjeto = RELAOBJ.tipobjasoc and "
    sql = sql & " OBJPRTY.estado = Estados.objid and "
    sql = sql & " RELAOBJ.codojasoc = OBJPRTY.codobjeto and "
    sql = sql & " RELAOBJ.tipobjasoc = OBJPRTY.tipobjeto "
    sql = sql & " and Objetos.codobjeto = " & nCodigo

  '  response.write "sql: " & sql & "<br>"

    set rsEstado = ofv.crearconsultaEx(Sql,cn,1,volver)
    if not rsEstado.eof then sEstado2 = rsEstado(0)
    call ofv.cerrarconsulta(rsEstado)

    if sEstado1 = 3 then


      sql = "Select objrem,Tipobjeto From Objetos Where codobjeto = " & nCodigo
      set rsCodigo = ofv.crearconsultaEx(Sql,cn,1,volver)
      if not rsCodigo.eof then
        if clng(rsCodigo(0)) <> 0 then
          sql = "Select Tipobjeto From Objetos Where codobjeto = " & rsCodigo(0)
          set rsCodigox = ofv.crearconsultaEx(Sql,cn,1,volver)
          if not rsCodigox.eof then xntipobj = rsCodigox(0)
'-----------------
          Segct = "S"

          Sql = "Select * From OBJSDK Where tipobjeto='"
          sql = sql & xntipobj & "' and codobjeto = " & rsCodigo(0)
          set rsx = ofv.crearconsultaEx(Sql,cn,1,parametros)
          if not rsx.eof then


          if Session("Usuario") <> rsx(7) then
             if rsx(3) then Segct = "N"
             call ofv.cerrarconsulta(rsx)
             sql = "Select * From OBJSDKDET Where tipobjeto='" & xntipobj
             sql = sql & "' and codobjeto = "  & rsCodigo(0)
             sql = sql & " and ((grpusr = 'USR-" & Session("Usuario") & "') "
             sql = sql & " or (grpusr = 'GRP-" & Session("Seggrp") & "')) "
           ''response.write Strsql & " sql " & "<br>"
             set rsx = ofv.crearconsultaEx(Sql,cn,1,parametros)
             do until rsx.eof

                select case rsx(3)
                   case "1"
                        Segct = "S"
                end select
              rsx.movenext
           loop
         end if
      end if
      call ofv.cerrarconsulta(rsx)

'-----------
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' ACTUALIZACIONES POR REEMPLAZO DE OBJETOS POR NUEVA VERSION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          if Segct = "S" then
             vFECHAHOY = clng(DATE) - nFecini
             sql = "Update Objprty Set estado = 4,fecbaja = '" & vFECHAHOY & "', "
             sql = sql & "elimino = '" & sUsuario & "' "
             sql = sql & "Where estado = 3 And codobjeto = " & rsCodigo(0)
             call ofv.crearconsultaEx(Sql,cn,1,volver)
             sql = "Update Objetos Set objrdo = " & nCodigo
             sql = sql & " Where codobjeto = " & rsCodigo(0)
             call ofv.crearconsultaEx(Sql,cn,1,volver)
             if cn.errors.count = 0 then

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' actualiza esquema asociados (arbol)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                sql = "Update Relaobj Set codobjmaster = " & nCodigo & " , tipobjmaster = '"
                sql = sql & rsCodigo(1) & "' Where codobjmaster = " & rsCodigo(0)
                sql = sql & " and tipobjmaster = '" & xntipobj & "' "
                call ofv.crearconsultaEx(Sql,cn,1,volver)




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' actualiza esquema linkeados
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                Sql = "Select * From objlnk Where tipobjetoL ='"
                sql = sql & xntipobj & "' and codobjetoL = " & rsCodigo(0)
                sql = sql & " and LTIPO = 'L'"
                set rsx2 = ofv.crearconsultaEx(Sql,cn,1,parametros)

           '''     response.write "-lee link ---" & Strsql & "<br>"

                DO WHILE not rsx2.eof

                   WIDEN3L = 0
                   WIDEN4  = 0
                   WIDEN4L = 0


'''''''''''''''''''''' fija la solapa '''''''''''''''''
                   WIDEN3L = RSX2(11)

'''''''''''''''''''''' fija el 1er nivel '''''''''''''''''

     SELECT CASE WIDEN3L

            CASE "2"   'CONSIDERACIONES

                     Strsql = "SELECT objid FROM OBJNWEDT WHERE relant = '" & RSX2(12) & "' and codobjeto = " & nCodigo
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        WIDEN4 = Xrs(0)
                     END IF
                     call ofv.cerrarconsulta(Xrs)


            CASE "5"   'PROCEDIMIENTOS

                     Strsql = "SELECT CODPROCED FROM objproced WHERE relant = '" & RSX2(12) & "' and codobjeto = " & nCodigo
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        WIDEN4 = Xrs(0)
                     END IF
                     call ofv.cerrarconsulta(Xrs)

            CASE "7"   'CONTROLES


                     Strsql = "SELECT objid FROM objcontrol WHERE relant = '" & RSX2(12) & "' and codobjeto = " & nCodigo
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        WIDEN4 = Xrs(0)
                     END IF
                     call ofv.cerrarconsulta(Xrs)


            CASE "8"   'ESQUEMAS

                     Strsql = "SELECT objid FROM objesquema WHERE relant = '" & RSX2(12) & "' and codobjeto = " & nCodigo
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        WIDEN4 = Xrs(0)
                     END IF
                     call ofv.cerrarconsulta(Xrs)


            CASE "9"   'TABLAS


                     Strsql = "SELECT objid FROM ObjTabDef WHERE relant = '" & RSX2(12) & "' and codobjeto = " & nCodigo
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        WIDEN4 = Xrs(0)
                     END IF
                     call ofv.cerrarconsulta(Xrs)


            CASE "11"   'CONTENIDO

                     Strsql = "SELECT objid FROM OBJESTR WHERE relant = '" & RSX2(12) & "' and codobjeto = " & nCodigo
                     set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                     if not Xrs.eof then
                        WIDEN4 = Xrs(0)
                     END IF
                     call ofv.cerrarconsulta(Xrs)



     END SELECT

          '''      response.write "---- lee anterior ----" & Strsql & "<br>"

'''''''''''''''''''''' fija el 2er nivel '''''''''''''''''

        IF  WIDEN3L = "9" THEN

            Strsql = "SELECT objid FROM ObjTabblk WHERE relant = '" & RSX2(13) & "' and codobjeto = " & nCodigo
            set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not Xrs.eof then
               WIDEN4L = Xrs(0)
            END IF
            call ofv.cerrarconsulta(Xrs)

        END IF


        IF  WIDEN3L = "11" THEN

            Strsql = "SELECT objid FROM OBJESTDET WHERE relant = '" & RSX2(13) & "' and codobjeto = " & nCodigo
            set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not Xrs.eof then
               WIDEN4L = Xrs(0)
            END IF
            call ofv.cerrarconsulta(Xrs)

        END IF

           ''''     response.write "---- lee anterior - 2do ---" & Strsql & "<br>"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




                   sql = "Update objlnk Set codobjetoL = " & nCodigo
                   sql = sql & ", tipobjetoL = '" & rsCodigo(1) & "' "
                   sql = sql & ", IDEN3L     = '" & WIDEN3L & "' "
                   sql = sql & ", IDEN4      = '" & WIDEN4 & "' "
                   sql = sql & ", IDEN4L     = '" & WIDEN4L & "' "
                   sql = sql & " Where OBJID = " & RSX2(0)
                   call ofv.crearconsultaEx(Sql,cn,1,volver)

             ''''   response.write "-------------- actualiza --" & sql & "<br>"



                RSX2.MOVENEXT
                LOOP




                sql = "Update objlnk Set codobjetoL = " & nCodigo & " , tipobjetoL = '"
                sql = sql & rsCodigo(1) & "' Where codobjetoL = " & rsCodigo(0)
                sql = sql & " and tipobjetoL = '" & xntipobj & "' "
                sql = sql & " and LTIPO <> 'L'"
                call ofv.crearconsultaEx(Sql,cn,1,volver)







''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

             else
                response.write ofv.ErrorDB(cn,volver)
             end if
          end if
        end if
      end if
      call ofv.cerrarconsulta(rsCodigo)
    elseif sEstado1 = 4 then
'''
        sql = "Select Tipobjeto From Objetos Where codobjeto = " & nCodigo
        set rsCodigox = ofv.crearconsultaEx(Sql,cn,1,volver)
        if not rsCodigox.eof then xntipobj = rsCodigox(0)
        Desde = 0
        Hasta = 0
        iz = 0
        Redim preserve aCodigo(iz)
        aCodigo(iz) = nCodigo
        Redim preserve atipobj(iz)
        atipobj(iz) = xntipobj

        do
          for nz = desde to hasta
            strsql = "Select codojasoc, tipobjasoc From Relaobj Where "
            strsql = strsql & " codobjmaster = " & aCodigo(iz)
            strsql = strsql & " and tipobjmaster = '" & atipobj(iz) & "' "
            set rsRel = ofv.crearconsultaEx(strsql,cn,1,volver)
            do until rsRel.eof
              iz = iz + 1
              Redim preserve aCodigo(iz)
              aCodigo(iz) = rsRel(0)
              Redim preserve atipobj(iz)
              atipobj(iz) = rsRel(1)
              rsRel.Movenext
            loop
            call ofv.cerrarconsulta(rsRel)
          next
          if hasta = iz then exit do
          desde = hasta + 1
          hasta = iz
        loop

        for iz = 0 to hasta
          vFECHAHOY = clng(DATE) - nFecini
          strsql = "Update Objprty Set estado = 4,fecbaja = '" & vFECHAHOY & "', "
          strsql = strsql & "elimino = '" & sUsuario & "' "
          strsql = strsql & "Where codobjeto = " & aCodigo(iz)
          call ofv.crearconsultaEx(StrSql,cn,1,volver)
        next
'''
    end if
  end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  set cn = ofv.conectar(ofv.strconn2)

' xxxxxxxxxxxxxxx ejecuta actualizacion: Update ---- xxxxxxxxxxxxxxxxxxxxxx
'reSponse.write "strsql:			www " & strsql
  if bEjecutar then

'    response.write "strsql: " & strsql
     call ofv.crearconsultaEx(StrSql,cn,1,volver)
  end if

'''' POR ERROR = 0 (PUEDE PROCEDER AL WKFL)
  if cn.errors.count = 0 then
     call ofv.cerrarconn(cn)
     set cn = ofv.conectar(ofv.strconn2)
     if tabla = "Objproced" then
        if request.form("codsector") = "nsnc" then
           call ofv.crearconsultaEx(StrSql,cn,1,volver)
        end if
     end if

'''' POR OPCION WKFLOW ("WFL")
     if  Ucase(sopcwfl) = "WFL" then

            XXXX = GETWKFMANAGER(nCodigo,STIPO,sestado1,tabla,"SMP")

     end if

    if sestado1 = 3 then
''''''      strsql = "Update Objprty Set aprobo = '" & sUsuario & "' Where codobjeto = " & nCodigo
    else

     if Ucase(tabla) <> "OBJWFLDET"  then
       strsql = "Update Objprty Set aprobo = '" & sUsuario & "' Where codobjeto = " & nCodigo
       call ofv.crearconsultaEx(StrSql,cn,1,volver)
     END IF

    end if

'response.write strsql
     call ofv.cerrarconn(cn)
    ' response.write " volver " & volver
' xxxxxxxxxxxxxxx Vuelve por OK al llamador ---- xxxxxxxxxxxxxxxxxxxxxx
'  response.write strsql
    response.redirect volver
   ' response.write volver
   else
'  response.write strsql
' xxxxxxxxxxxxxxx sale por Error ---- xxxxxxxxxxxxxxxxxxxxxx
'response.write strsql
   response.write ofv.ErrorDB(cn,volver)
     call ofv.cerrarconn(cn)
  end if


end if
end if




%>