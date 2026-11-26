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
  Scompe = request.querystring("compe")
  SOPCION = request.querystring("OPCION")

  if scompe <> "" then
     session("compe") = scompe
  else
     scompe = session("compe")
  end if
   
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



  sHTML = ofv.FormHeader(session("FormN"))







  Strsql = "Select * From tipocompetencia where objid = 0" & scompe
'' response.write strsql  
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   
    
  
   if  not rs.eof then


   Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","20%","","","","TIPO DE COMPETENCIA:","si")
   sInput = RS(1)
   Celda = Celda & ofv.gencelda("","utI ","","80%","","","",sInput,"si")

   Fila = ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")


   shtml =shtml & ofv.GenTabla("","AAA","","70%   ","","0","","0","0","","",fila,"si") & "<BR><BR>"


   end if
  







  'Encabezados
  Formulario = ofv.Encabezados("#","Descripcion")




''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA competencias - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN





        strsql = "Insert into tipocompeDET "
        strsql = strsql & " (COMPETENCIA,descripcion,SECUEN)"
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & SCOMPE & "','NUEVO DETALLE DE COMPETENCIA','0'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)




 END IF

  xTrColor = "#FFFFFF" 


  Strsql = "Select * From TipocompeDET WHERE COMPETENCIA = " & SCOMPE & " Order By SECUEN, descripcion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=TipocompeDET&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(3),sHref)

    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'id
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"15","dtn","")
    Celda = Celda & ofv.gencelda("","UT VALIGN='TOP' ","","2%","","","",sInput,"si")


    'Descripcion
   ' sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"100","dt","")

    sInput = "<font class=at><textarea class=dt name=" & rs(2).name & "  id=" & rs(2).name & " rows=3 cols=80 "
    if bModificar  then
       sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 450) {alert('El texto es demasiado"
       sInput = sInput & " extenso'); return false;} else {"
       sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    else
       sInput = sInput & " readonly "
    end if
    sInput = sInput & " >" &  trim(rs(2)&"")  &"</textarea></font>"
    


    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")




    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0" 
       ELSE
          xTrColor = "#FFFFFF" 
       END IF

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=TipocompeDET&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","1","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "tipocompefunDET.asp?opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)

  end if



  str1 = "Select descripcion From TipocompeDET WHERE COMPETENCIA = " & SCOMPE & " Order By descripcion "    ' orden pagina
  str2 = "Select descripcion From TipocompeDET WHERE COMPETENCIA = " & SCOMPE & " Order By descripcion "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  
  'consulta

 ' sAccion = "TipodscptoCons.asp?" & request.querystring
 ' Celda = ofv.BotonConsultar(sAccion)

  sAccion = "TipoCOMPEFUN.asp"
  Celda = ofv.BotonVolver(sAccion)
  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>