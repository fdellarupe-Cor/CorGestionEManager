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
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)


  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("Descripcion","DEGRADACION","OBSERVACIONES","COLOR","&nbsp;")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN

        strsql = "Insert into ORACTIMPACTO "
        strsql = strsql & " (VALOR,DESCRI,IMPDESDE,IMPHASTA,OBS1,COLOR)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('0','NUEVO NIVEL IMPACTO','0','0','OBSERVACION','#BBBBBB')"

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF


  WKSETTABLA = "ORACTIMPACTO"

  Strsql = "Select * From " & WKSETTABLA & " Order By IMPDESDE"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(2),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


    'Descripcion
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"8","dtN","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")



    'OBSERVACIONES
   ''' sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm,"30","dt","")



       nCol = 40



       if  not  isnull(rs(5)) then
           wdescri = rs(5)
       end if

       if  len(wdescri) = 0 then wdescri = " "

       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' "
       t3b = t3b & "name=" & rs(5).name & " cols=" & nCol & " rows=3 "
       if  bModificar then
           t3b = t3b & " onblur=" & sForm & ".submit() "
           t3b = t3b & " onkeypress=" & chr(34)
           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
       else
           t3b = t3b & " readonly >"
       end if
       xdatos =  wdescri
       t3b    = t3b & xdatos & "</textarea></font>"
       SINPUT = t3b


    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'color
    scomK = "#00DD00" & chr(9) & "#F7EEDD" & chr(9) & "#F7DD77" & chr(9) & "#FF6666" & chr(9) & "#993333" & chr(9) & "#BBBBBB" & chr(9) & "#888888"
    scomK = scomK & chr(9) & "#555555" & chr(9)  & "#333333" & chr(9)  & "#000000" & chr(9) & "#FF8C00"

    scomD = "VERDE CLARO" & chr(9) & "BEIGE" & chr(9) & "AMARILLO" & chr(9)
    scomD = scomD & "ROJO CLARO" & chr(9) & "ROJO OSCURO" & chr(9) & "GRIS CLARO" & chr(9)
    scomD = scomD & "GRIS" & chr(9) & "GRIS INTERMEDIO" & chr(9) & "GRIS OSCURO" & chr(9)
    scomD = scomD & "NEGRO" & CHR(9) & "NARANJA"
    sInput = ofv.gencomboFX(sForm,rs(6).name,rs(6),scomK,scomD," ")
    Celda = Celda & ofv.gencelda("","UT","","20%","","","",sInput,"si")

    sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:" & RS(6) & ";' ","CENTER","5%","","","",sInput,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "RGACTIMPACTO.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA & " Order By IMPDESDE "    ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & " Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "RGACTIMPACTOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>