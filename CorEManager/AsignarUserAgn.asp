<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  SMODAL     = request.querystring("MODAL")
  id         = request.querystring("id")
  opc=request.querystring("opc")
  stipo=request.querystring("tipo") ' SI/NO --> SI ES SI UPDATE TODA LA AGENDA AL NUEVO USUARIO
  SIDUSER=Request.querystring("IDUSER")
  cantregmover=request.querystring("cantregmover")

  Volver = request.querystring("Volver")


 APLCHECK()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if opc="UPD"  then
     set cn = ofv.conectar(ofv.strconn2)
     xusu=request.form("trcusrreg")
     strsql="update orAGENDA set orAGENDA.RESPONSABLE='"&xusu&"'"

  '   FECact  = CLNG(DATE()) - clng(cdate("01-01-1900"))



      if stipo="NO" then 'Solo  una entrada en la agenda
       strsql=strsql & " where OBJID='" & id & "'"
      ELSE
        'strsql=strsql & " where RESPONSABLE='" & SIDUSER & "'"
        strsql=strsql & " from (" & session("filtrosql") & ") jj where (jj.objid=oragenda.objid) "
      end if

    ' RESPONSE.WRITE  strsql
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)




     'response.write strsql
     call ofv.cerrarconn(cn)
     sAccion = "ORREASIGNAAGENDAEX.ASP?cantregmover=" & cantregmover & "&tipo=" & stipo &"&IDUSER=" & SIDUSER
     '"orbasEriesgoSadd.asp?modal=" & smodal & "&codtracer=" & scodtracer & "&cantregmover=0"
     response.redirect saccion
end if

     sHTML = ofv.FormHeader("Agregar Usuario")

     SHTML = SHTML & GenBuscaUsr("N","N","N","S",CHR(9),CHR(9),"GRP-","USR-")

     IF SESSION("FBUSR") <> "NO" THEN

        set cn = ofv.conectar(ofv.strconn2)




        f1a = "AsignaruserAGN.asp?opc=UPD"
        f1a = "AsignaruserAGN.asp?opc=UPD&id=" & id &"&cantregmover=" & cantregmover & "&tipo=" & stipo &"&IDUSER=" & SIDUSER
        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")



       SELECT CASE SESSION ("XGUMASTER")
        CASE "U"
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & id & "' NAME='CODIGO' >"
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & REPLACE(SESSION("FBUSR"),"USR-","")  & "' NAME='trcusrreg' >"



        CASE "S"
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & id & "' NAME='CODIGO' >"
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR")  & "' NAME='" & RS(3).NAME & "' >"

       END SELECT

      ELSE

        sAccion ="ORREASIGNAAGENDAEX.ASP?cantregmover=" & cantregmover & "&tipo=" & stipo &"&IDUSER=" & SIDUSER

        sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")
      END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






  'Botones Aceptar y Cancelar
  sAccion = "ORREASIGNAAGENDAEX.ASP?cantregmover=" & cantregmover & "&tipo=" & stipo &"&IDUSER=" & SIDUSER
  'modal=" & smodal & "&codtracer=" & scodtracer & "&cantregmover=0"
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>