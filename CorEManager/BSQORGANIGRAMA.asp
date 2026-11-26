<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  BSQ=request.querystring("BSQ")
  id=request.querystring("NID")
  
  
  IF SESSION("BSQ")="" THEN
     SESSION("BSQ")=BSQ
  END IF
  
  IF SESSION("IDRET")="" THEN
   SESSION("IDRET")=ID
  ELSE
  ID=SESSION("IDRET")
  END IF

     session("RGPROC") = RGPROC

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")

	'   sHTML = ofv.FormHeader(session("FormN"))

	 '----------------------------------------------

'	   shtml = shtml & CorPnlFnc("N","")

 ' sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"


     SELECT CASE BSQ
     CASE "PORDEPENDENCIA"
       XXBSQ= GenBuscaUsrEx("N","N","N","N","S","N","N","N","N","N","","","","","","","","","","")
     CASE "PORPUESTO"
       XXBSQ= GenBuscaUsrEx("N","N","N","N","N","S","N","N","N","N","","","","","","","","","","")
     CASE "PORNOMBRE"
       XXBSQ= GenBuscaUsrEx("N","N","N","n","N","N","S","N","N","N","","","","","","","","","","")
     END SELECT
     
     
     SHTML = SHTML & XXBSQ 'GenBuscaUsrEx("N","N","N","S","S","S","N","N","N","N","",CHR(9),CHR(9),CHR(9),"","","","","","")

     IF SESSION("FBUSR") <> "NO" THEN

        set cn = ofv.conectar(ofv.strconn2)
        
        SELECT CASE SESSION("BSQ")
        CASE "PORDEPENDENCIA"
        f1a ="OrganigramaPuestosgraf.asp?id=" & SESSION("FBUSR")
        CASE "PORPUESTO"
        f1a ="GENPUESTOFUNCION.asp?PUESTO=" & SESSION("FBUSR")
        
        CASE "PORNOMBRE"
        Set cn = ofv.conectar(ofv.strconn0)
        sql="select idpst from puestorel where cod='" & SESSION("FBUSR") &"'"
         Set rspx = ofv.crearconsultaEx(sql, cn, 1, "")
         
         if not rspx.eof then 
         idp=rspx(0)
         else
           IDP=id
         end if  
         
         call ofv.cerraRconsulta(rspx)
         
          Call ofv.cerrarconn(cn)
        IF   SESSION("BSQ")="PORNOMBRE" THEN
             f1a ="verinfoUSR.asp?id="&SESSION("FBUSR") & "&IDP="&IDP&"&bsq=USUARIO&VERFIRMAS=SI"
             f1a= f1a & "&XGUFILTRO=" & request.form("XGUFILTRO")
              f1a= f1a & "&XGUITEMCD=" & request.form("XGUITEMCD") 
        ELSE
        f1a ="OrganigramaPuestosgraf.asp?id=" & IDP
        END IF
        
        
        END SELECT
        
        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")
     
        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='XXX' >"
        
       
      ELSE
   
        sAccion ="OrganigramaPuestosgraf.asp?id=" & id
       
  

        sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")


      END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






  'Botones Aceptar y Cancelar
  sAccion ="OrganigramaPuestosgraf.asp?id=" & SESSION("IDRET") 

  
 
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self" ) 
  if  SESSION("FBUSR") ="NO" THEN
  sHTML = sHTML &  "<SCRIPT>"
  sHTML = sHTML &  "var ele = document.getElementById('AcepFRM');"
  sHTML = sHTML &  "ele.style.visibility = 'hidden';"
  sHTML = sHTML &  "</SCRIPT>"
  END IF
   
'--------------------------------------------------------------------------------
'  sHTML = sHTML & "</div>"
'
'  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
''  Fila = ofv.GenRow("","","","","","","",Celda,"si")
'  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")
'
''  'Estructura del Indice
'
'  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,"")
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>" '&  
  'response.write request.form
end if

response.write sHTML

%>