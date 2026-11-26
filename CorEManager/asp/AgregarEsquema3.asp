<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

sTipo = request.querystring("Tipo")
sNombre = request.querystring("Nombre")
sUsuario = session("Usuario")
sElemento = request.querystring("Elemento")
n = instr(sElemento,"-")
if n > 0 then
  nCodMaster = left(sElemento,n-1)
  sTipoMaster = mid(sElemento,n+1)
  n = instr(sTipoMaster,"-")
  sTipoMaster = left(sTipoMaster,n-2)
end if

SPTIPO = request.querystring("SPTIPO")

'response.write " elemento " & selemento & " tipomaster " & stipomaster

on error resume next

set cn = ofv.conectar(ofv.strconn2)
nFecini = 0



if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

strsql = "select * from Objetos"
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Insert into Objetos (Tipobjeto"

For Each x In Request.form
  strsql = strsql & "," & x
Next


strsql = strsql & ") Values ('" & sTipo & "'"

For Each yy In Request.form
x = lcase(yy)
IF x = "aplic" THEN WKFAPLIC = Request.Form(x)
  TipoValido = ofv.ValidarTipo(rs(x),Request.Form(x))
  if TipoValido = False and lcase(rs(x).name) <> "codiden" then exit for
  ValorNulo = ofv.EsNulo(rs(x),Request.Form(x))
  if ValorNulo = True and rs(x).name <> "codiden" then exit for
  s = ofv.queesx(rs(x).type)
  if rs(x).type = 200 then s = "'"
  if lcase(rs(x).name) = "codiden" then
       strsql = strsql & ",'" & cstr(session.sessionid) & "'"
       TipoValido = True
       ValorNulo = False
  else
       strsql = strsql & "," & s & UCase(Request.Form(x)) & s
  end if
Next


volver = "../vuelve.asp"
if TipoValido = False or ValorNulo = True then
  on error resume next
  response.write ofv.ErrorTipo(volver,Request.Form(x),rs(nx).name)
  if err.number <> 0 then response.write ofv.ErrorTipo(volver,"","")
else
  strsql = strsql & ")"
 ' response.write "primer insert " & strsql & "<br>"
  call ofv.crearconsultaEx(StrSql,cn,1,volver)

  if cn.errors.count = 0 then
    call ofv.cerrarconsulta(rs)

    strsql = "Select codobjeto From Objetos Where codiden = '" & cstr(session.sessionid) & "'"

    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then nCodAsoc = rs(0)

    call ofv.cerrarconsulta(rs)

    if request.form("codiden") = "" then
       ncodiden = "SD"
    else
       ncodiden =  request.form("codiden")
    end if

    strsql = "Update Objetos set codiden = '" & ncodiden & "', indobj = 10 "
    strsql = strsql & " where codiden = '" & cstr(session.sessionid) & "'"
    call ofv.crearconsultaEx(StrSql,cn,1,volver)


    Strsql = "insert into OBJdistPrty  (codobjeto,tipobjeto,emisor,asunto,texto,importancia,aviso,dias) values ("
    Strsql = Strsql & nCodAsoc & ",'" & stipo & "','" & session("usuario") & "','Activacion de <ESQNOR>',"
    Strsql = Strsql & "'<ESQNOR> esta vigente a partir de la fecha',1,'NO','7')"
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
'     response.write "2do insert " & strsql & "<br>"

    sNIVE = "D"

    StrSql = "Select TIPOCPT From Tipobjeto "
    StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & stipo & "'"
    Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If  Not rsIcono.EOF Then
        sNIVE  = rsIcono(0)
    End If
    Call OFV.cerrarconsulta(rsIcono)


    IF SESSION("EMPMUN") = "S" THEN

    Strsql = "insert into OBJdist  (codobjeto,tipobjeto,GRPUSR,SECTOR) values ("
    Strsql = Strsql & nCodAsoc & ",'" & stipo & "','TODOS','" & session("UN") & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
'    response.write "3 insert " & strsql & "<br>"
    END IF

    strsql = "Insert into Relaobj (Codobjmaster,Tipobjmaster,Codojasoc,Tipobjasoc) "
    strsql = strsql & "Values (" & nCodMaster & ",'" & sTipomaster & "',"
    strsql = strsql & nCodAsoc & ",'" & sTipo & "')"
   ' response.write "4 insert " & strsql & "<br>"
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
nEstado = 1
    if cn.errors.count = 0 then

       IF sNIVE = "D" THEN
          strsql = "Select objid From Estados Where estado = '" & UCASE(session("menu")) & "'"

' response.write "estado " & strsql & "<br>"
         set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
         if not rs.eof then nEstado = rs(0)
       ELSE
            nEstado = 3
       END IF

       if WKFAPLIC = "RG" then nEstado = 3
       if WKFAPLIC = "PL" then nEstado = 1

'      dFechaNula = "1/1/1900"
      dFechaNula = clng(cdate("01-01-1900")) - nFecini
      dHoy = clng(Date) - nFecini
      strsql = "Insert into Objprty (codobjeto, tipobjeto, estado, fecemis, fecvigen, "
      strsql = strsql & "fecbaja, fecvto, fecacta,fecpublic, autor) Values (" & nCodAsoc & ",'" & sTipo & "',"
      strsql = strsql & nEstado & ",'" & DHoy & "','" & DFechaNula & "','"
      strsql = strsql & DFechaNula & "','" & DFechaNula & "','" & dFechaNula & "','" & dFechaNula & "','" & sUsuario & "')"

'     response.write "5 insert " & strsql & "<br>"

      call ofv.crearconsultaEx(StrSql,cn,1,volver)
      if cn.errors.count = 0 then
         strsql = "Insert into OBJSDK (codobjeto, tipobjeto, segct,segmo,segco,segce,segown "
         strsql = strsql & ") Values (" & nCodAsoc & ",'" & sTipo & "',0,0,0,0,'"
         strsql = strsql & Session("Usuario") & "')"
        ' response.write "6 insert " & strsql & "<br>"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)
         if cn.errors.count = 0 then
            strsql = "Insert into OBJWFL (codobjeto, tipobjeto)"
            strsql = strsql & " Values (" & nCodAsoc & ",'" & sTipo & "')"
           '  response.write "7 insert " & strsql & "<br>"
            call ofv.crearconsultaEx(StrSql,cn,1,volver)
            if cn.errors.count = 0 then

'''''''''''' SOLO PARA SUBPROCESOS CON VINCULACION EXPLICITA (ACTIVIDADES ADMINISTRADAS EN FORMA MANUAL)
               IF SPTIPO = "CR" or SPTIPO = "AI" THEN


                  GETFECH()

                  WKFECHA = XLFECDAT - clng(cdate("01-01-1900"))
                  WKHORA  = XLHORDAT
                  WKSECTOR = 0

                  strsql = "insert into OBJSECTOR  ("
                  strsql = strsql & "codobjeto,tipobjeto,codSECTOR,MODI,"
                  strsql = strsql & "TIPMODI,MODAUTOR,MODFECHA,MODHORA"
                  strsql = strsql & ") values ("
                  strsql = strsql & nCodAsoc & ",'" & sTipo & "','0','1',"
                  strsql = strsql & "'A','" & SESSION("USUARIO") & "','" & WKFECHA & "','" & WKHORA & "'"
                  strsql = strsql & ")"
                 ' response.write "8 insert " & strsql & "<br>"
                  call ofv.crearconsultaEx(StrSql,cn,1,volver)

                  strsql = "SELECT OBJID FROM OBJSECTOR "
                  strsql = strsql & " WHERE CODOBJETO = " & nCodAsoc & " "
                  set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                  if NOT rsN.eof then
                     WKSECTOR = rsN(0)
                  END IF


                  sevento = 0
                  sukey = GTkey(nCodAsoc)

                  strsql = "insert into objproced  ("
                  strsql = strsql & "codobjeto,tipobjeto,codproced,codeven,"
                  strsql = strsql & "frecuencia,tipproced,CODSECTOR,MODI,"
                  strsql = strsql & "TIPMODI,MODAUTOR,MODFECHA,MODHORA"
                  strsql = strsql & ") values ("
                  strsql = strsql & nCodAsoc & ",'" & sTipo & "','" & sukey & "','" & sevento & "',"
                  strsql = strsql & "'DIARIA','" & wkaplic & "','" & WKSECTOR & "','1',"
                  strsql = strsql & "'A','" & SESSION("USUARIO") & "','" & WKFECHA & "','" & WKHORA & "'"
                  strsql = strsql & ")"
               '   response.write "9 insert " & strsql & "<br>"
                  call ofv.crearconsultaEx(StrSql,cn,1,volver)




                  strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
                  strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values ('" & nCodMaster & "','" & sTipomaster & "',"
                  strsql = strsql & "'" & nCodMaster & "','" & nCodMaster & "',"
                  strsql = strsql & "'" & nCodAsoc & "','" & sTipo & "','','','LRGCR','5','" & sukey & "','0'"
                  strsql = strsql & ")"
               '   response.write "10 insert " & strsql & "<br>"
                  call ofv.crearconsultaEx(StrSql,cn,1,volver)

                  Strsql = "SELECT * From OBJLNK where iden1 = '" & nCodMaster & "' and iden2 = '" & nCodMaster & "' and LTIPO = 'LRGCR' "
                  set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                  if not rsN.eof then
                     wlnkid = rsN(0)
                  else
                     wlnkid = 0
                     wlink = 0
                  end if
                  call ofv.cerrarconsulta(rsN)

                  if clng(wlnkid) > 0 then
                     strsql = "update OBJETOS set rellink = " & wlnkid
                     strsql = strsql & " where  codobjeto = 0" & nCodMaster
                     call ofv.crearconsultaEx(StrSql,cn,1,volver)
                  end if



               END IF

'''''''''''' SOLO PARA OBJETOS SOPORTE DEL PLAN DE CONTINGENCIAS
               IF SPTIPO = "PL" THEN


                  GETFECH()

                  WKFECHA = XLFECDAT - clng(cdate("01-01-1900"))
                  WKHORA  = XLHORDAT
                  WKSECTOR = 0


                  strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, iden3, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
                  strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values ('" & nCodMaster & "','" & sTipomaster & "',"
                  strsql = strsql & "'" & nCodMaster & "','" & nCodMaster & "','" & SESSION("ORPERIODO") & "',"
                  strsql = strsql & "'" & nCodAsoc & "','" & sTipo & "','','','LRGPL','0','0','0'"
                  strsql = strsql & ")"
               '   response.write "11 insert " & strsql & "<br>"
                  call ofv.crearconsultaEx(StrSql,cn,1,volver)

                  Strsql = "SELECT * From OBJLNK where iden1 = '" & nCodMaster & "' and "
                  strsql = strsql & " iden2 = '" & nCodMaster & "' and "
                  strsql = strsql & " iden3 = '" & SESSION("ORPERIODO") & "' and "
                  strsql = strsql & "LTIPO = 'LRGPL' "
                  set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                  if not rsN.eof then
                     wlnkid = rsN(0)
                  else
                     wlnkid = 0
                     wlink = 0
                  end if
                  call ofv.cerrarconsulta(rsN)

                  if clng(wlnkid) > 0 then
                     strsql = "update OBJETOS set rellink = " & wlnkid
                     strsql = strsql & " where  codobjeto = 0" & nCodMaster
                     call ofv.crearconsultaEx(StrSql,cn,1,volver)
                  end if



               END IF


             '   response.redirect volver
            else
                response.write ofv.ErrorDB(cn,volver)
            end if
        else
           response.write ofv.ErrorDB(cn,volver)
        end if
      else
        response.write ofv.ErrorDB(cn,volver)
      end if
    else
      response.write ofv.ErrorDB(cn,volver)
    end if
  else
    response.write ofv.ErrorDB(cn,volver)
  end if
end if

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>