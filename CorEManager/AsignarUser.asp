<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  SCODTRACER = request.querystring("codtracer")
  SMODAL     = request.querystring("MODAL")
  id         = request.querystring("id")
  opc=request.querystring("opc")


  Volver = request.querystring("Volver")


 APLCHECK()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if opc="UPD"  then
     set cn = ofv.conectar(ofv.strconn2)
     xusu=request.form("trcusrreg")
     strsql="update ortracer set TRCMARCAACTP='A',trcusrreg='" & xusu & "'"

     FECact  = CLNG(DATE()) - clng(cdate("01-01-1900"))

     strsql= strsql & " ,TRCFECACTP=" & FECact

     'if session("APLMDCLI") = "CFA" THEN
      xx=getGroupsUser(xusu)
      grupow= split(session("SGU_grupos"),chr(9))
      strsql= strsql & " ,TRCGRUPOWNER='"& grupow(0) &"'"
     'end if

     strsql=strsql & " where objid=0" & id
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     STRSQL="SELECT TRCCODTRACERP FROM ORTRACER WHERE  objid=0" & ID
     SET RSXP=ofv.crearconsultaEx(StrSql,cn,1,volver)

      IF NOT RSXP.EOF THEN

       IF (RSXP(0) &"") <>"" THEN

       strsql="update ortracer set TRCMARCAACTH='A' "
	   strsql= strsql & " ,TRCFECACTH=" & FECact
	   strsql=strsql & " where CODTRACER='" & RSXP(0) &"'"
       call ofv.crearconsultaEx(StrSql,cn,1,volver)
       END IF
      END IF

      CALL OFV.CERRARCONSULTA(RSXP)



     'response.write strsql
     call ofv.cerrarconn(cn)
     sAccion = "orbasEriesgoSadd.asp?modal=" & smodal & "&codtracer=" & scodtracer & "&cantregmover=0"
     response.redirect saccion
end if

     sHTML = ofv.FormHeader("Agregar Usuario")

     SHTML = SHTML & GenBuscaUsr("N","S","N","S",CHR(9),CHR(9),"GRP-","USR-")

     IF SESSION("FBUSR") <> "NO" THEN

        set cn = ofv.conectar(ofv.strconn2)




        f1a = "Asignaruser.asp?opc=UPD"
        f1a = "Asignaruser.asp?opc=UPD&modal=" & smodal & "&codtracer=" & scodtracer & "&cantregmover=0&id=" & id
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

        sAccion ="orbaseriesgoadd.asp?id=" & id

        sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")
      END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






  'Botones Aceptar y Cancelar
  sAccion = "orbasEriesgoSadd.asp?modal=" & smodal & "&codtracer=" & scodtracer & "&cantregmover=0"
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>