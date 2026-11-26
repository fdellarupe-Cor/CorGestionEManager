<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%




response.expires=0


nCodigo = request.querystring("Codobjeto")
sTipo = request.querystring("Tipobjeto")

sdato = request.form("dato")

volver = "EDobjnove.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&opc=SI"

set cn = ofv.conectar(ofv.strconn2)



if not sdato  then 
     strsql = "delete from objnoved where  codobjeto =" & nCodigo & " and tipobjeto ='" & stipo & "' " 
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)
     response.redirect volver 
else

' busca perfil del documento si no existe lo genera




        strsql = "Insert into Objnoved (codobjeto,tipobjeto,texto,dias,tipodist"
        strsql = strsql & ") values (" & nCodigo & ",'" & sTipo & "','<esqnor> esta vigente a partir de <fecvig>',10,1)" 


    call ofv.crearconsultaEx(StrSql,cn,1,parametros)
 



' Fin generacion de novedad del documento

     response.redirect volver 
end if  



     call ofv.cerrarconn(cn)





%>