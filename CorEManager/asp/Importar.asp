<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

nProced = request.querystring("Codproced")
nCodigo = request.querystring("Codobjeto")
sTipo = request.querystring("Tipobjeto")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
nNewProced = request.form(1)
volver = volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

set cn = ofv.conectar(ofv.strconn2)

vFECHAHOY = clng(DATE) - clng(cdate("01-01-1900"))
xxxx      = GETFECH()
vHORAHOY  = XLHORDAT
vUSUARIO  = SESSION("USUARIO")

strsql = "Select * From Objaccion Where codproced = '" & nNewProced & "'"
strsql = strsql & " and (OBJACCION.modi = 0 or (OBJACCION.modi = 1 and OBJACCION.tipmodi <> 'E')) "
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
ind = 0
do until rs.eof
  ind = ind + 1
  sukey = GTkey(nCodigo & ind)
''  response.write sukey & "<br>"

  if clng(rs(19)) > 0 then
        Strsql = "SELECT * From OBJLNK where objid = " & rs(19)
        set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs2.eof then


          ' strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL "
          ' strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "',"
          ' strsql = strsql & "'" & nProced & "','" & sukey & "',"
          ' strsql = strsql & rs2(5) & ",'" & rs2(6) & "'"
          ' strsql = strsql & ")"


           strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2,iden3, codobjetoL, tipobjetoL,LTipo,Lurl,LurlDesc "
           strsql = strsql & ",iden3l,iden4,iden4l) Values (" & nCodigo & ",'" & sTipo & "',"
           strsql = strsql & "'" & nProced & "','" & sukey & "','" & rs2("iden3") & "',"
           strsql = strsql & rs2("codobjetoL") & ",'" & rs2("tipobjetoL") & "','" & rs2("LTipo") & "','" & rs2("Lurl") & "','" & rs2("LurlDesc") & "'"
           strsql = strsql & ",'" & rs2("iden3l") & "','" & rs2("iden4") & "','" & rs2("iden4l") & "'"
           strsql = strsql & ")"


           call ofv.crearconsultaEx(StrSql,cn,1,volver)
''  response.write strsql & " 1<br>"

        end if
        call ofv.cerrarconsulta(rs2)

        Strsql = "SELECT objid From OBJLNK where iden2 = '" & sukey & "' "
        set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs2.eof then
           wlink = rs2(0)
        else
           wlink = 0
        end if
        call ofv.cerrarconsulta(rs2)

        strsql = "Insert into Objaccion (Codobjeto,Tipobjeto,Codproced,Tipoaccion,modo,codcomp,"
        strsql = strsql & "tipocomp,sectorcont,perfilcont,secuen,referencia,tiempo,copias,"
        strsql = strsql & "modi,tipmodi,referencia2,referencia3,codaccion,rellink,reflink,"
        strsql = strsql & "MODAUTOR,MODFECHA,MODHORA) Values("
        strsql = strsql & nCodigo & ",'" & sTipo & "','" & nProced & "'," & rs(4) & ",'" & rs(5) & "',"
        strsql = strsql & rs(6) & ",'" & rs(7) & "'," & rs(8) & ",'" & rs(9) & "'," & rs(10) & ",'"
        strsql = strsql & rs(11) & "'," & rs(12) & "," & rs(13) & ",1,'A'"
        strsql = strsql & ",'" & rs(16)  & "','" & rs(17)& "','" & sukey & "'," & wlink & ",'" & rs(20) & "'"
        strsql = strsql & ",'" & VUSUARIO & "','" & vFECHAHOY & "','" & vHORAHOY & "')"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

''  response.write "<br>" & strsql & " 1<br>"

   else

        strsql = "Insert into Objaccion (Codobjeto,Tipobjeto,Codproced,Tipoaccion,modo,codcomp,"
        strsql = strsql & "tipocomp,sectorcont,perfilcont,secuen,referencia,tiempo,copias,"
        strsql = strsql & "modi,tipmodi,referencia2,referencia3,codaccion,rellink,reflink,"
        strsql = strsql & "MODAUTOR,MODFECHA,MODHORA) Values("
        strsql = strsql & nCodigo & ",'" & sTipo & "','" & nProced & "'," & rs(4) & ",'" & rs(5) & "',"
        strsql = strsql & rs(6) & ",'" & rs(7) & "'," & rs(8) & ",'" & rs(9) & "'," & rs(10) & ",'"
        strsql = strsql & rs(11) & "'," & rs(12) & "," & rs(13) & ",1,'A'"
        strsql = strsql & ",'" & rs(16)  & "','" & rs(17)& "','" & sukey & "',0,''"
        strsql = strsql & ",'" & VUSUARIO & "','" & vFECHAHOY & "','" & vHORAHOY & "')"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

''  response.write strsql & " 2<br>"

   end if


'  strsql = "Select objid From Objaccion Where codproced = " & nProced
'  set rsNuevo = ofv.crearconsulta2(StrSql,cn)
'  rsNuevo.movelast
'  nUltimo = rsNuevo(0)
'  call ofv.cerrarconsulta(rsNuevo)

  strsql = "Select * From Objaccioncontrol Where idaccion = '" & rs(18) & "'"
  strsql = strsql & " and (OBJACCIONcontrol.modi = 0 or (OBJACCIONcontrol.modi = 1 and OBJACCIONcontrol.tipmodi <> 'E')) "
  set rsAlt = ofv.crearconsultaEx(StrSql,cn,1,volver)
  do until rsAlt.eof
    strsql = "Insert Into Objaccioncontrol (codproced,accion,idaccion,secuencia,referencia,"
    strsql = strsql & "continuador,modi,tipmodi,"
    strsql = strsql & "MODAUTOR,MODFECHA,MODHORA) Values("
    strsql = strsql & "'" & nProced & "'," & rsAlt(2) & ",'" & sukey & "',"
    strsql = strsql & rsAlt(4) & ",'" & rsAlt(5) & "','" & rsAlt(6) & "',1,'A'"
    strsql = strsql & ",'" & VUSUARIO & "','" & vFECHAHOY & "','" & vHORAHOY & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
    rsAlt.movenext
  loop
  call ofv.cerrarconsulta(rsAlt)
  rs.movenext
loop
call ofv.cerrarconsulta(rs)

if cn.errors.count = 0 then
  response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if

call ofv.cerrarconn(cn)

%>