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
  sopcion  = request.querystring("opcion") & ""
  evento1=request.querystring("evento1")
  lotemap=request.querystring("lotemap")
  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezad
  IF SOPCION <> "MOD" THEN
   Formulario = ofv.Encabezados("CUENTA","CARACTER","EVENTO NIV.1","EVENTO NIV.2","PRODUCTO","PROPIEDADES")
  ELSE
  Formulario = ofv.Encabezados("CUENTA","CARACTER","EVENTO NIV.1","EVENTO NIV.2","PRODUCTO")
  END IF

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
 ' response.write request.querystring
  set cn = ofv.conectar(ofv.strconn2)


'''''''''''''''''''''''''''''''''''''''
'' INCORPORO LOTE NUEVO


  IF SOPCION=  "SI" THEN

     sfechai = clng(date()) - clng(cdate("01-01-1900"))

        strsql = "Insert into ORLOTEMAP "
        strsql = strsql & " (CUENTA,CARACTER,TEVENTO1,TEVENTO2,PRODUCTO"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'CUENTA','0','0','0','0'"
        strsql = strsql & ")"
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)
        SOPCION=""
END IF
SELECT CASE  sopcion

CASE "MOD"


    WKSETTABLA = "ORLOTEMAP"

     Strsql = "Select tevento1 From ORLOTEMAP a where "

     strsql = strsql & "   a.objid=0" & lotemap

     set rsw = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    if not rsw.eof then
   ' response.write cstr(evento1 & "") &"-----" & cstr(rs("tevento1")&"")
     if cstr(evento1 & "") <> cstr(rsw("tevento1")&"") then
      call ofv.cerrarconsulta(rsw)
      strsql = "update ORLOTEMAP set tevento2=0   where objid=0" & lotemap
    '  response.write strsql
      set rsw = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     end if
    else
     call ofv.cerrarconsulta(rsw)
    end if

     Strsql = "Select * From " & WKSETTABLA & " WHERE OBJID=" & LOTEMAP
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     posI = 0
     Do While (Not rs.EOF) and (posi < CantRegMostrarx)
       sForm = "form" & posi

       wkestado = rs(5) & ""

       'Eliminar
      ' if bSuprimir AND wkestado = "0" then
       if bSuprimir  then
         sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
         sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
         Celda = ofv.BotonEliminar(rs(1),sHref)
       else
         Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
       end if

       'CUENTA
       SINPUT = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"20","dt","")
       Celda = Celda & ofv.gencelda("","CC","","15%","","","",sInput,"si")

        'CARACTER
          Strsql = "Select * From ORTIPOCARACTER ORDER BY DESCRI "
          sInput = ofv.GenerarCombo(sForm,RS(2).NAME,RS(2).value,Strsql,cn,0,1,"")
       Celda = Celda & ofv.gencelda("","CC","","15%","","","",sInput,"si")

       'EVENTO1
          Strsql = "Select * From ORTIPOEVENTO ORDER BY DESCRI"
          sInput = ofv.GenerarCombo(sForm,RS(3).NAME,RS(3).value,Strsql,cn,0,1,"")
       Celda = Celda & ofv.gencelda("","CC","","15%","","","",sInput,"si")

       'EVENTO2
          Strsql = "Select * From ORTIPOCATEGO  where tipoevento="& RS(3).value &" order BY DESCRI"

          sInput = ofv.GenerarCombo(sForm,RS(4).NAME ,RS(4).value,Strsql,cn,0,2,"")
          Celda = Celda & ofv.gencelda("","CC","","15%","","","",sInput,"si")

       'PRODUCTO
          Strsql = "Select * From TIPOPRODUCTO ORDER BY DESCRI "
          sInput = ofv.GenerarCombo(sForm,RS(5).NAME,RS(5).value,Strsql,cn,0,1,"")
          Celda = Celda & ofv.gencelda("","CC","","15%","","","",sInput,"si")



       Fila = ofv.genrow("","","","","","","",Celda,"si")

       sAccion = ""
       if bModificar then
         sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
         sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta &"_evento1=" & rs("tevento1") &"_lotemap=" & rs(0) & "_OPCION=MOD"
       end if
       Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

       posi = posi + 1
       rs.Movenext
  Loop


  CASE ELSE



  WKSETTABLA = "ORLOTEMAP"


 STRSQL =" Select A.*,B.DESCRI AS BCARACTER,C.DESCRI AS CEVENTO1,D.DESCRI AS DEVENTO2 ,E.DESCRI AS EPRODUCTO From ORLOTEMAP A "
 STRSQL = STRSQL & "  LEFT JOIN ORTIPOCARACTER B ON(A.CARACTER=B.OBJID)"
 STRSQL = STRSQL & "  LEFT JOIN ORTIPOEVENTO C ON(A.TEVENTO1=C.OBJID)"
 STRSQL = STRSQL &  " LEFT JOIN ORTIPOCATEGO D ON(A.TEVENTO2=D.OBJID)"
 STRSQL = STRSQL &  " LEFT JOIN TIPOPRODUCTO E ON(A.PRODUCTO=E.OBJID) ORDER BY A.OBJID DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF)
    sForm = "form" & posi

    wkestado = rs(5) & ""

    'Eliminar
   ' if bSuprimir AND wkestado = "0" then
    if bSuprimir  then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'CUENTA
    SINPUT = RS(1)
    Celda = Celda & ofv.gencelda("","CC","","15%","","","",sInput,"si")

     'CARACTER

       sInput = RS("BCARACTER")
    Celda = Celda & ofv.gencelda("","CC","","15%","","","",sInput,"si")

    'EVENTO1

       sInput =RS("CEVENTO1")
    Celda = Celda & ofv.gencelda("","CC","","15%","","","",sInput,"si")

    'EVENTO2

       sInput = RS("DEVENTO2")
       Celda = Celda & ofv.gencelda("","CC","","15%","","","",sInput,"si")

    'PRODUCTO

       sInput =RS("EPRODUCTO")
       Celda = Celda & ofv.gencelda("","CC","","15%","","","",sInput,"si")

      sImagenP = "&#0232;"
	  sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
      sinput = "<a class='ft1n' " & sFuncAnc & " href='ORINTERFASESMAP.asp?evento1=" & rs("tevento1") &"&opcion=MOD&LOTEMAP=" & rs(0) & "&CantRegMover=" & CantRegMoverx
      sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
      sinput = sinput  & "' title='Propiedades de la cuenta " & rs(1) & "' >" & sImagenP & "</a>"




    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta &"_evento1=" & rs("tevento1") &"_lotemap=" & rs(0)
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
 END SELECT
  sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar AND sopcion <>"MOD" then
    sAccion = "ORINTERFASESMAP.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  IF sopcion <>"MOD" THEN
  str1 = "Select CUENTA From " & WKSETTABLA & " Order By OBJID DESC "    ' orden pagina
  str2 = "Select CUENTA From " & WKSETTABLA & " Order By CUENTA "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
  sAccion = "ORINTERFASES.asp"
  Celda = ofv.BotonVolver(sAccion)
  ELSE
  BUSCAR=""
  sAccion = "ORINTERFASESMAP.asp?CANTREGMOVER=" & CantRegMoverx &"&CANTREGMOSTRAR=8"
  Celda = ofv.BotonVolver(sAccion)
  END IF




  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>