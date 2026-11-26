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
''  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")


 if sopcion = "SI" then
    GENESTADO()

 end if



    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>&#0232;</FONT></B>"

    celdas = ""



  'Encabezados
  formulario = formulario & ofv.GenCelda("","ttt colspan=5","","","","","","cuentas de presupuesto/asignacion","si")
  formulario = formulario & ofv.Encabezados("COD.","orden","descripcion","cod.interno","subcuentas","ESTADO")


  StrSql = "Select * From PROYcuentas order by orden"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rs(0)


       sc = "ut"

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=PROYcuentas&volver=" & vuelta & "&Eliminar=NSNC"
      Celdas = ofv.BotonEliminar(rs(1),sHref)
    else
      Celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

 '   celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")



       sInput = ofv.GenerarInput("",rs(0),"readonly","","5","dtn","")
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

       if bModificar then
          sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sform,"3","dt","")
       else
          sInput = ofv.GenerarInput("",rs(3),"readonly","","3","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")




       if bModificar then
          sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sform,"60","dt","")
       else
          sInput = ofv.GenerarInput("",rs(1),"readonly","","60","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


       if bModificar then
          sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sform,"30","dt","")
       else
          sInput = ofv.GenerarInput("",rs(2),"readonly","","30","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")



                       scomK = "1" & chr(9) & "2"
                       scomD = "ACTIVO" & chr(9) & "CERRADO"


                       if bModificar then
                          sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD," ")

                       else
                          if rs(2) = 1 then sInput = "ACTIVO"
                          if rs(2) = 2 then sInput = "CERRADO"
                       end if



       celdas = celdas & ofv.GenCelda("",sc,"CENTER","20%","","","",sInput,"si")


       sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";'  href=PRycuentasdet.asp?TIPO=" & rs(0)
       sinput2 = sinput2  & " TITLE='SUBcuentas DE " & rs(1) & "'  ><b>" & sImagenC & "</b></a>" 

       celdas = celdas & ofv.GenCelda("",sc,"CENTER","","","","",sInput2,"si")




    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=PROYcuentas"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
    rs.Movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "PRYcuentas.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""
  

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML

FUNCTION GENESTADO()

  XCODORDEN = 0


  StrSql = "Select MAX(ORDEN) From PROYcuentas "
  set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsX.eof then XCODORDEN = RSX(0)
  call ofv.cerrarconsulta(rsX)

  XCODORDEN = XCODORDEN + 10


    strsql = "insert into PROYcuentas (DESCRIPCION,CODINT,ORDEN,ESTADO) values ('NUEVA CUENTA','SC','" & XCODORDEN & "','1')"
    
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION



%>