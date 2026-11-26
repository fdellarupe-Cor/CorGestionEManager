<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  sMODAL        = request.querystring("MODAL")
  sCODTRACER    = request.querystring("CODTRACER")
  sOPCION       = request.querystring("OPCION")


  sHTML = ofv.FormHeader(session("FormN"))


  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_CODTRACER=" & sCODTRACER & "_MODAL=" & sMODAL

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


'''''''''''''''''''''''''''''''''''''''
'' INCORPORO contabilizacion para evento BASE DE EVENTOS DE PERDIDA

  if sOPCION = "SI" then

     SORDEN = 10


     Strsql = "select MAX(ORDEN) FROM ORTRACERCONT WHERE CODTRACER = '" & SCODTRACER & "' "

'''  response.write strsql & "<br>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rs.eof then
        SORDEN = RS(0) + 10
     END IF


     sCODTRACERCONT = GTkey("500")

        strsql = "Insert into ORTRACERCONT "
        strsql = strsql & " (CODTRACERCONT,CODTRACER,ORDEN,CTACTBLE, DESCRIPCION, MONTOIMP"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & sCODTRACERCONT & "','" & sCODTRACER & "','" & SORDEN & "',"
        strsql = strsql &  "'CUENTA','DESCRIPCION DE LA IMPUTACION','0'"
        strsql = strsql & ")"
''  response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if

'''''''''''''''''''''''''''''''''''''''


   Celda = ""

   wkproccod = "&nbsp;"
    
   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","BASE DE EVENTOS DE RIESGO - REGISTRO DE CONTABILIZACION ","si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")  
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")




  'Encabezados
  Formulario = ""

    Celda = ""


    Formulario = ""

  Strsql = "select * FROM ORTRACERCONT WHERE CODTRACER = '" & SCODTRACER & "' ORDER BY ORDEN, OBJID DESC   "

''  response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     Celda = ""
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO EXISTEN REGISTROS DE CONTABILIZACION PARA EL EVENTO","si")
     Fila = ofv.GenRow("","","","","","","",Celda,"si")
     Tabla = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

  sHTML = sHTML & "<CENTER><BR><BR><BR><BR>" & Tabla & "</CENTER>"


  ELSE
  posI = 0



    sInput = "&nbsp;"
    Celda = ofv.gencelda("","UT colspan=3","","","","","",sInput,"si")
    Formulario = ofv.genrow("","AAA bgcolor='#FFFFFF' ","","","25","","",Celda,"si")

  Do While (Not rs.EOF)
    sForm = "form" & posi
    Filax  = ""





    Celda = ""
    'REGISTRO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;' ","","2%","","","","#","si")
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;' ","","38%","","","","CUENTA","si")

    'OBSERVACIONES
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","60%","","","","DESCRIPCION","si")
    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda = ""

    sInput = ofv.GenerarInput(rs(3).name,rs(3).value,"TEXT",sForm,"3","dtN","")
    Celda = Celda & ofv.gencelda("","UTI ","","","","","",sInput,"si")

    sInput = ofv.GenerarInput(rs(4).name,rs(4).value,"TEXT",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","UTI ","","","","","",sInput,"si")

      nCol = 80

      if  len(wdescri) = 0 then wdescri = "Sin Descripcion" 

      if  not  isnull(rs(5)) then 
          wdescri = rs(5)
      end if 
  
      t3b = "<font class=at><textarea class=dt "
      t3b = t3b & "name=" & rs(5).name & " cols=" & nCol & " rows=4 "
      if  bModificar then
          t3b = t3b & " onblur=" & sForm & ".submit() "
          t3b = t3b & " onkeypress=" & chr(34)
          t3b = t3b & " if (this.value.length > 180) {alert('El texto es demasiado"
          t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
      else
          t3b = t3b & " readonly >"
      end if
      xdatos =  wdescri
      t3b    = t3b & xdatos & "</textarea></font>"
    SINPUT = t3b
    Celda = Celda & ofv.gencelda("","UTI ROWSPAN=3 ","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ""
    'REGISTRO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;' ","","2%","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;' ","","40%","","","","MONTO","si")
    Filax  = Filax & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ""

    Celda = Celda & ofv.gencelda("","UTI ","","","","","","&nbsp;","si")

    sInput = ofv.GenerarInput(rs(6).name,rs(6).value,"TEXT",sForm,"30","dtN","")
    Celda = Celda & ofv.gencelda("","UTI ","","","","","",sInput,"si")
    Filax  = Filax & ofv.genrow("","","","","","","",Celda,"si")


    sInput = "&nbsp;"
    Celda = ofv.gencelda("","UT colspan=3","","","","","",sInput,"si")
    Filax  = Filax & ofv.genrow("","AAA bgcolor='#FFFFFF' ","","","25","","",Celda,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=ORTRACERCONT&volver=" & vuelta
    end if
      Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")



    posi = posi + 1
    rs.Movenext
  Loop
  END IF
  sHTML = sHTML & ofv.GenTabla("","AAA BGCOLOR='#ddddff' ","","98% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "ORBASERIESGOSCONT.asp?CODTRACER=" & SCODTRACER & "&MODAL=" & SMODAL & "&OPCION=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if
  Buscar  = ""
  Celda   = ""

  sAccion = "ORBASERIESGOSADD.asp?CODTRACER=" & SCODTRACER & "&MODAL=" & SMODAL
  Celda = ofv.BotonVolver(sAccion)
  
  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)


  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>