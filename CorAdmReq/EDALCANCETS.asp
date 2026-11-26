<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

'  IF session("ModMain") <> "WF" THEN
'     bAgregar = FALSE
'     bModificar = FALSE
'     bSuprimir = FALSE
'  END IF

 sopcion = request.querystring("opcion")

     xSUBTEMA = 0
     xTEMA = 0

  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xSUBTEMA = RS(9) 
     IF XSUBTEMA > 0 THEN 
        SQL2 = "SELECT * "
        SQL2 = SQL2 & " FROM csbuscasubtema where "
        SQL2 = SQL2 & " codPRODSERV = " & xSUBTEMA
        set rs = ofv.crearconsultaEx(SQL2,cn3,1,parametros)
        if NOT rs.eof then
           xTEMA = RS(3)
        END IF
     END IF

  END IF

 if sopcion = "R" THEN

    sTema = REQUEST.FORM("codtema")
    sSubtema = 0
 ELSEIF SOPCION = "SI" THEN
    sSubtema = REQUEST.FORM("codprodserv")
    StrSql = "UPDATE REQUERIMIENTO SET TEMASUBTEMA = " & sSubtema & " where codigo = '" & session("NCliente") & "' "
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
    RESPONSE.REDIRECT "EDALCANCE.ASP"

 ELSE

   sTema    = xTEMA
   sSubtema = xSUBTEMA
 END IF 


    SQL1 = "SELECT * From Tema Order By descri"


    SQL2 = "SELECT * "
    SQL2 = SQL2 & " FROM csbuscasubtema where "
    SQL2 = SQL2 & " codtema = " & sTema
    SQL2 = SQL2 & " Order By descri"





    sHTML = sHTML & "<center>"
    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;ALCANCE DEL REQUERIMIENTO","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")




    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","","","40%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","right","10%","","","","Tema:&nbsp;&nbsp;","si")
    sInput = ofv.GenerarCombo("Subconsulta1","codtema",sTema,sql1,cn3,0,1,"")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = "EDALCANCETS.asp?OPCION=R"
    Formulario = ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila =  ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","right","","","","","Subtema:&nbsp;&nbsp;","si")
    sInput = ofv.GenerarCombo("Subconsulta2","codprodserv",sSubtema,sql2,cn3,0,1,"")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = "EDALCANCETS.asp?OPCION=SI"
    Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")



    FilaS =  filaS & ofv.GenRow("","","","","","","",formulario,"si")
    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


    sHTML = sHTML & ofv.GenTabla(""," ","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"


  Agregar = ""
  Buscar  = ""
  Celda   = ""


     sAccion = "EDALCANCE.asp"
     Celda = ofv.BotonVolver(sAccion)
 

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","","no",Agregar,Buscar,celda)



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>