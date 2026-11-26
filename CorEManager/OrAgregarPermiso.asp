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
  sCod = request.querystring("Cod")
  sCIRCUITO    = request.querystring("CIRCUITO")
  Volver = request.querystring("Volver")
  Volver = Volver & "?Cod=" & sCod & "_CIRCUITO=" & SCIRCUITO
  sxopc1 = ""


     session("RGPROC") =""

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")

     SHTML = SHTML & GenBuscaUsr("N","N","S","S",CHR(9),CHR(9),"GRP-","USR-")

     IF SESSION("FBUSR") <> "NO" THEN

        set cn = ofv.conectar(ofv.strconn2)



           strsql = "Select * From orusrlst "

           WKSPTABLA = "orusrlst"




       set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


        f1a = "asp/AgregarEN.asp?Tabla=" & WKSPTABLA
        f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
        f1a = f1a & "&Tipobjeto=" & sTipo

        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

         sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & nCodigo & "' NAME='" & RS(1).NAME & "' >"
         if session("scod") <> 1 then
          sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(3).NAME & "' >"
         else
         sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(2).NAME & "' >"
         end if
      '  sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & sCod & "' NAME='" & RS(2).NAME & "' >"

      ELSE

        sAccion = "orasgusracc.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod  & "&CIRCUITO=" & SCIRCUITO & "&volver=edobjsdk.asp"

        sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")


      END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






  'Botones Aceptar y Cancelar
  sAccion = "orasgusracc.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=edobjsdk.asp" & "&CIRCUITO=" & SCIRCUITO
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>