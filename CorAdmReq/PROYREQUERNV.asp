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

 IF SOPCION = "CP" THEN 

     StrSql = "Select * From PROYECTO where codigo = '" & session("NCliente") & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rs.eof then

    sHTML = ofv.formHeader(" ")

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"
  
    celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;PROYECTO ASOCIADO AL REQUERIMIENTO","si")
    filas  = ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","VT STYLE='BORDER-STYLE:ridge;BORDER-WIDTH:2;border-color:#f7aa22;' COLSPAN=3","CENTER","","50","","","<BIG><B>CONFIRMA LA ASOCIACION DE UN PROYECTO AL REQUERIMIENTO ??</B></BIG>","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     sinput = "<a  class=tt href=PROYREQUERnv.asp?OPCION=CPX"
     sinput = sinput  & " ><b>&nbsp;CONFIRMA&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

     sinput = "<a  class=tt href=CONSULTAR.asp"
     sinput = sinput  & " ><b>&nbsp;CANCELA&nbsp;</b></a>" 
     celdas = CELDAS & ofv.GenCelda("","tt","","10%","","","",sInput,"si")

    celdas = CELDAS & ofv.GenCelda("","","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","1","0","","",filas,"si")

  sHTML = sHTML  & "</body></html>"
  response.write sHTML
        
     ELSE

    saccion = "editarPRY.asp?codigo=" & session("NCliente") & "&tipo=PRY"

    sHTML = ofv.formHeader(" ")

    sHTML = sHTML & "<form name='dsph' action='" & saccion & "' method='post' target='_top' ><input type='hidden' value='S'></form>"

    sHTML = sHTML & "<script language=javascript> document.dsph.submit(); </script>"

    sHTML = sHTML  & "</body></html>"

    response.write sHTML

    '    RESPONSE.REDIRECT "conspry.ASP"
     END IF 


 END IF

 IF SOPCION = "CPX" THEN 

       prycod = OWKF.GENPROY(session("NCliente"))

     '  response.write prycod & "<BR>"

       sanio = datepart("yyyy",date())
       smes  = datepart("m",date())
       smesC = smes

       Strsql = "Select * From meses WHERE mes = " & smes
       set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
       if not rs.eof then
          smesC = RS(2)
       END IF 


       Session("VPLAN")    = 0 

       OWKF.GENPROP()

    '   response.write prycod & "-2<BR>"

       STRSQL = "UPDATE PROYPRTY SET CONF = 1 WHERE CODIGO = '" & session("NCliente") & "'  "
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

       OWKF.GENINITPRY()

    ''   YXETP = OWKF.GENPRYETP(Session("NCliente"), Session("VPLAN"), "10", "INCIDENCIAS REALIZADAS", "0", "0")

     '  response.write prycod & "-3<BR>"


       StrSql = "UPDATE PROYECTO SET CUENTA = 2, MODAL = 0  WHERE CODIGO = '" & prycod & "'"
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

       wdescri = "ADMINISTRACION DE INCIDENCIAS"

       XXXX = OWKF.PRYLNCH(prycod,wdescri,"S")

       StrSql = "DELETE From PROYTAREAS WHERE CODIGO = '" & Session("NCliente") & "' "
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

       XXXX =  GENSECTOROR() 

       XXXX =  GENUSRCONT()


 


    saccion = "editarPRY.asp?codigo=" & session("NCliente") & "&tipo=PRY"

    sHTML = ofv.formHeader(" ")

    sHTML = sHTML & "<form name='dsph' action='" & saccion & "' method='post' target='_top' ><input type='hidden' value='S'></form>"

    sHTML = sHTML & "<script language=javascript> document.dsph.submit(); </script>"

    sHTML = sHTML  & "</body></html>"

    response.write sHTML

    ''    RESPONSE.REDIRECT "conspry.ASP"



 END IF




end if

' PROYETPTSKVPLAN

%>