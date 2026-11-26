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
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sOPCION       = request.querystring("OPCION")
  SCODINDI       = request.querystring("INDI")
  ID            =request.querystring("ID")





  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



'''''''''''''''''''''''''''''''''''''''
'' INCORPORO INDICADOR

  if sopcion = "SIz" then

        STRSQL="SELECT * FROM ORINDICADORES WHERE OBJID=0" & ID
        SET RSIND= ofv.crearconsultaEx(StrSql,cn,1,volver)
        if RSIND("tipo")=999 then
         tipo=2

        else
         tipo=1

        end if
        sfechai = clng(date()) - clng(cdate("01-01-1900"))

        strsql = "Insert into ORINDICADORES "
        strsql = strsql & " (DESINDICADOR,ACCESO,LSTUSR,TIPO,DESCRIPCION,FECHACREACION,estado,usuario,FORMULA,CANTVAR"
        strsql = strsql & " ) Values ('"&RSIND("DESINDICADOR")&"',1,1,"&tipo&","
        strsql = strsql  &"'"& RSIND("DESCRIPCION")&"',"& sfechai
        strsql = strsql & ",1,'" & session("usuario") & "',"& RSIND("FORMULA")&"," & clng("0"&RSIND("CANTVAR"))&")"
        'response.write strsql
        call ofv.crearconsultaEx(StrSql,cn,1,volver)
        RESPONSE.REDIRECT "ORCARGAUSRINDI.ASP"

  end if



     Strsql = "select * FROM orindicadores WHERE TIPO in (0,999) and estado=0"

   Celda = ""

   wkproccod = "&nbsp;"

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","TABLA DE INDICADORES","si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")




  'Encabezados
  Formulario = ""

    Celda = ""


    Formulario = ""



 response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

    Celda = ofv.gencelda("","vT  style='border-style:SOLID;border-width:1;border-color:#f93333;' ","center","","","","","<big>NO HAY CARGADOS INDICADORES</big>","si")
    Formulario = ofv.genrow("","","","","60","","",Celda,"si")

  ELSE

    Celda = ""

    Celda = Celda & ofv.gencelda("","VTI COLSPAN=2 STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","INDICADOR","si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","DESCRIPCION","si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","ASIGNAR INDICADOR","si")

    Formulario = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  posI = 0

  Do While (Not rs.EOF)
    sForm = "form" & posi

    Celda = ""
    'Tipo
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","5","","","",sInput,"si")

    'indicador
    sInput = RS(1) 'ofv.GenerarInput(rs(1).name,rs(1),"text READONLY",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'descripcion


   nCol = 60



   if  not  isnull(rs(3)) then
	   wdescri = rs(3)
   end if

   if  len(wdescri) = 0 then wdescri = " "

   t3b = "<font class=at><textarea readonly class=dt style='font-size:7pt;' "
   t3b = t3b & "name=" & rs(3).name & " cols=" & nCol & " rows=2 "



   xdatos =  wdescri
   t3b    = t3b & xdatos & "</textarea></font>"
   SINPUT = "<font class=at>" & wdescri &"</font>"  't3b



  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")


  'ASIGNAR

  	sImagenP = "&#0232;"
  	sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
  	sinput = "<a class='ft1n' " & sFuncAnc & " href='ORCARGAINDICADORES.asp?OPCION=SI&ID=" & rs(0) & "&CantRegMover=" & CantRegMoverx
  	sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
  	sinput = sinput  & "' title='Propiedades de " & rs(1) & "' >" & sImagenP & "</a>"
    Celda = Celda & ofv.gencelda("","ft","LEFT","","","","",sInput,"si")



    Filax  = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=ORINDICADORES&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")


    posi = posi + 1
    rs.Movenext
  Loop
  END IF
  sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar  then
    sAccion = "ORINDICADORES.asp?opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","",Buscar,celda)

  call ofv.cerrarconsulta(rs)




  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>