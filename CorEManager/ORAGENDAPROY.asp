<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  set cn = ofv.conectar(ofv.strconn6)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


 sopcion = request.querystring("opcion")
 sACCION = request.querystring("ACCION")
 sAGENDA = request.querystring("AGENDA")

 IF SOPCION = "CP" THEN 

    session("NCliente") = sAGENDA


     StrSql = "Select * From PROYECTO where codigo = '" & session("NCliente") & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rs.eof then

    sHTML = ofv.formHeader(" ")

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=corflaticon SIZE=4>4</FONT>"
  
    celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;PLANIFICACION DEL PLAN DE MITIGACION","si")
    filas  = ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","VT STYLE='BORDER-STYLE:ridge;BORDER-WIDTH:2;border-color:#f7aa22;' COLSPAN=3","CENTER","","50","","","<BIG><B>CONFIRMA LA ASOCIACION DE UN PROYECTO AL PLAN DE MITIGACION ??</B></BIG>","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     sinput = "<a  class=tt href=ORAGENDAPROY.asp?OPCION=CPX&AGENDA=" & SAGENDA
     sinput = sinput  & " ><b>&nbsp;CONFIRMA&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

     sinput = "&nbsp;" 
     celdas = CELDAS & ofv.GenCelda("","Ut","","10%","","","",sInput,"si")

    celdas = CELDAS & ofv.GenCelda("","","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","1","0","","",filas,"si")

  sHTML = sHTML  & "</body></html>"
  response.write sHTML
        
     ELSE

    saccion = "../CORADMREQ/editarPRY.asp?codigo=" & session("NCliente") & "&tipo=PRY"

    sHTML = ofv.formHeader(" ")

    sHTML = sHTML & "<form name='dsph' action='" & saccion & "' method='post' target='_top' ><input type='hidden' value='S'></form>"

    sHTML = sHTML & "<script language=javascript> document.dsph.submit(); </script>"

    sHTML = sHTML  & "</body></html>"

    response.write sHTML

    '    RESPONSE.REDIRECT "conspry.ASP"
     END IF 


 END IF

 IF SOPCION = "CPX" THEN 

        XXXX = OWKF.GENPROY(session("NCliente"))

    saccion = "../CORADMREQ/editarPRY.asp?codigo=" & session("NCliente") & "&tipo=PRY"

    sHTML = ofv.formHeader(" ")

    sHTML = sHTML & "<form name='dsph' action='" & saccion & "' method='post' target='_top' ><input type='hidden' value='S'></form>"

    sHTML = sHTML & "<script language=javascript> document.dsph.submit(); </script>"

    sHTML = sHTML  & "</body></html>"

    response.write sHTML

    ''    RESPONSE.REDIRECT "conspry.ASP"



 END IF

 IF SOPCION = "CT" THEN 

     StrSql = "Select * From PROYETPTSK where PROYTSKOBJ = '" & session("NCliente") & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rs.eof then

        RESPONSE.REDIRECT "PROYREQUERTSKASOC.ASP"

     END IF 
     
        RESPONSE.REDIRECT "PROYREQUERTSK.ASP"

 END IF


end if

' PROYETPTSKVPLAN

%>