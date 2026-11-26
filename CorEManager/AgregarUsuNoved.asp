<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  id = request.querystring("id")

  Volver = request.querystring("Volver")
'  Volver = Volver & "?Cod=" & sCod & "_CIRCUITO=" & SCIRCUITO
  sxopc1 = ""
'response.write "STRING: " & request.querystring & "<br>" & "FORM: " & request.form

     session("RGPROC") =""
'response.write request.querystring & "<br>" & request.form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")
     select case session("APLMDCLI")

     CASE "GP"
         wwGenBuscaUsr= GenBuscaUsrEx("S","N","N","S","N","N","N","N","N","N","","","","","","","","","","")
     CASE ELSE
         wwGenBuscaUsr= GenBuscaUsr("N","S","N","S",CHR(9),CHR(9),"GRP-","USR-")
     END SELECT
     SHTML = SHTML &  wwGenBuscaUsr
     'response.write SESSION("FBUSR")
     IF SESSION("FBUSR") <> "NO" THEN

        set cn = ofv.conectar(ofv.strconn2)



           strsql = "Select * From novegenpardist "

       WKSPTABLA = "NOVEGENPARDIST"




       set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


        f1a = "asp/AgregarEN.asp?Tabla=" & WKSPTABLA
        f1a = f1a & "&volver=" & volver
        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")



       SELECT CASE SESSION ("XGUMASTER")
        CASE "U"
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & id & "' NAME='CODIGO' >"
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & REPLACE(SESSION("FBUSR"),"USR-","")  & "' NAME='" & RS(2).NAME & "' >"
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='0' NAME='" & RS(3).NAME & "' >"
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='0' NAME='" & RS(4).NAME & "' >"
        case "E"
           sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & id & "' NAME='CODIGO' >"
		   sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(3).NAME & "' >"
           sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='TODOS' NAME='" & RS(2).NAME & "' >"
           sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='0' NAME='" & RS(4).NAME & "' >"
        CASE "S"
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & id & "' NAME='CODIGO' >"
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR")  & "' NAME='" & RS(3).NAME & "' >"
		  sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='SECT-' NAME='" & RS(2).NAME & "' >"
		  sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='0' NAME='" & RS(4).NAME & "' >"
       END SELECT

      ELSE

        sAccion ="ADnovedADESDist.asp?id=" & id ' "qqqorasgusracc.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&i=" & sCod  & "&CIRCUITO=" & SCIRCUITO & "&volver="

        sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")
      END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





XGUFILTRO = request.form("XGUFILTRO")

  'Botones Aceptar y Cancelar
'if XGUFILTRO <> "TEFND" and XGUFILTRO <> "TE" then
  sAccion = "ADnovedADESDist.asp?id=" & id
  zbtn= ofv.AgregarBoton("AgreArchb",sAccion,"_self")

'end if
if SESSION("XGUFILTRO")="TEFNDs" or SESSION("XGUFILTRO")="TE" then zbtn=replace(zbtn,"Aceptar","Aceptar' style='display:none'")

sHTML = sHTML & zbtn
sAccion = "ADnovedADESDist.asp?id=" & id
volver = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","",cantregmoverx,cantregmostrarx,"no","no",posi,"","","no","","",volver)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>" & XGUMASTER
end if
'response.write
response.write sHTML

%>