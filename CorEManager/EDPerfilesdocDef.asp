<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%




response.expires=0


nCodigo = request.querystring("Codobjeto")
sTipo = request.querystring("Tipobjeto")

sdato = request.form("dato")

volver = "EDperfilesdocProp.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&opc=SI"

set cn = ofv.conectar(ofv.strconn2)



if sdato  then 
     strsql = "delete from objperfildoc where  codobjeto =" & nCodigo & " and tipobjeto ='" & stipo & "' " 
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)
     response.redirect volver 
else

' busca perfil del documento si no existe lo genera



     Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
     set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs1.eof then 
        sPerfil = rs1(3)
     end if
     call ofv.cerrarconsulta(rs1)

     Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
     set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs1.eof then 

        strsql = "Insert into Objperfildoc (codobjeto,tipobjeto"
        strsql2 = ") values (" & nCodigo & ",'" & sTipo & "'"

    for ii = 0 to (rs1.fields.count-1)
      strsql = strsql & "," & rs1(ii).name
      s = ofv.QueEsx(rs1(ii).type)
      if rs1(ii).type > 100 and  rs1(ii).type < 300 then s = "'"

      S = "'"
      sinput = rs1(ii)

      if ucase(sinput) = "VERDADERO" then sinput = "1"
      if ucase(sinput) = "FALSO" then sinput = "0"
      if ucase(sinput) = "TRUE" then sinput = "1"
      if ucase(sinput) = "FALSE" then sinput = "0"
 

      if rs1(ii).type = 131 then sinput = replace(rs1(ii),",",".")

   
      strsql2 = strsql2 & "," & s & sinput & s 

      wkcampo = rs1(ii).name

      if ucase(wkcampo) = "OBJORD15" THEN EXIT FOR

    next
  
    strsql = strsql & strsql2 & ") "

 '' response.write "sql " & strsql & "<br>"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)
 
     end if
     call ofv.cerrarconsulta(rs1)


' Fin generacion del perfil del documento

     response.redirect volver 
end if  



     call ofv.cerrarconn(cn)





%>