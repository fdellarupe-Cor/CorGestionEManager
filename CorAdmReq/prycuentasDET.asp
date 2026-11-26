<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))


  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons2" & chr(34) & "," & chr(34) & "scrollbars=yes,resizable=yes,menubar=no,toolbar=no,height=450,width=800" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"


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
  XTIPO = REQUEST.QUERYSTRING("TIPO")

  IF XTIPO <> "" THEN
     SESSION("TIPO") = XTIPO
  ELSE
     XTIPO = SESSION("TIPO")
  END IF

 if sopcion = "SI" then
    GENESTADO()

 end if

     XTIPON = XTIPO

  StrSql = "Select * From PROYCUENTAS WHERE ID = 0" & XTIPO
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     XTIPON = RS(1)
  END IF


    celdas = ""


    Xhoy = CLNG(DATE()) - clng(cdate("01-01-1900"))

    StrSql = "UPDATE PROYCUENTASDET SET FINICIO = " & Xhoy & " WHERE CODCUENTA = 0" & XTIPO
    StrSql = StrSql & " AND FINICIO = 0 AND ACTIVA = 1"
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

    StrSql = "UPDATE PROYCUENTASDET SET FFINAL = " & Xhoy & " WHERE CODCUENTA = 0" & XTIPO
    StrSql = StrSql & " AND FFINAL = 0 AND ACTIVA = 2"
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)


    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>&#0232;</FONT></B>"




  'Encabezados
  formulario = formulario & ofv.GenCelda("","ttt colspan=8","","","","","","SUBCUENTAS&nbsp;PARA&nbsp;:&nbsp;" & XTIPON,"si")
  formulario = formulario & ofv.Encabezados("COD.","ORDEN","DESCRIPCION","INICIO","FIN","PLAN","CONS","ESTADO","DETALLE")


  StrSql = "Select * From PROYCUENTASDET WHERE CODCUENTA = 0" & XTIPO
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rs(0)


       sc = "uT"


    'Eliminar
    if bSuprimir AND cstr(RS(8) & "") = "0" AND RS(6) <> 1 then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=PROYCUENTASDET&volver=" & vuelta & "&Eliminar=NSNC"
      Celdas = ofv.BotonEliminar(rs(1),sHref)
    else
      Celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


       sInput = rs(0)
       celdas = celdas & ofv.GenCelda("",sc,"RIGHT","1%","","","",sInput,"si")


       if bModificar then
          sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sform,"3","dtN","")
       else
          sInput = ofv.GenerarInput("",rs(3),"readonly","","3","dtN","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","1%","","","",sInput,"si")



       if bModificar then
          sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sform,"50","dt","")
       else
          sInput = ofv.GenerarInput("",rs(2),"readonly","","50","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","10%","","","",sInput,"si")


       XESTADO = CLNG("0" & RS(6))

       SELECT CASE XESTADO

              CASE 0
                   sInput = "---"
                   celdas = celdas & ofv.GenCelda("",sc,"CENTER","5%","","","",sInput,"si")

                   sInput = "---"
                   celdas = celdas & ofv.GenCelda("",sc,"CENTER","5%","","","",sInput,"si")



              CASE 1

                   sInput = CDATE(CLNG("0" & rs(4))  + clng(cdate("01-01-1900")) )
                   celdas = celdas & ofv.GenCelda("",sc,"CENTER","5%","","","",sInput,"si")

                   sInput = "---"
                   celdas = celdas & ofv.GenCelda("",sc,"CENTER","5%","","","",sInput,"si")



              CASE ELSE


                   sInput = CDATE(CLNG("0" & rs(4))  + clng(cdate("01-01-1900")) )
                   celdas = celdas & ofv.GenCelda("",sc,"CENTER","5%","","","",sInput,"si")

                   sInput = CDATE(CLNG("0" & rs(5))  + clng(cdate("01-01-1900")) )
                   celdas = celdas & ofv.GenCelda("",sc,"CENTER","5%","","","",sInput,"si")


        END SELECT 






       if bModificar then
          sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sform,"5","dtN","")
       else
          sInput = ofv.GenerarInput("",rs(7),"readonly","","5","dtN","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"RIGHT","5%","","","",sInput,"si")

       sInput = rs(8)
       celdas = celdas & ofv.GenCelda("",sc,"RIGHT","5%","","","",sInput,"si")





       SELECT CASE XESTADO

              CASE 0

                   IF  NOT (RS(7) > 0) THEN

                       sInput = "NO ACTIVO"
                       celdas = celdas & ofv.GenCelda("",sc,"CENTER","","","","",sInput,"si")


                   ELSE

                       scomK = "0" & chr(9) & "1"
                       scomD = "NO&nbsp;ACTIVO" & chr(9) & "ACTIVO"


                       if bModificar then
                          sInput = ofv.gencomboFX(sForm,rs(6).name,rs(6),scomK,scomD," ")

                       else
                          if rs(6) = 0 then sInput = "NO&nbsp;ACTIVO"
                          if rs(6) = 1 then sInput = "ACTIVO"
                          if rs(6) = 2 then sInput = "CERRADO"
                       end if



                       celdas = celdas & ofv.GenCelda("",sc,"CENTER","20%","","","",sInput,"si")


                   END IF  



              CASE 1

                       scomK = ""
                       scomD = ""


                   IF  (cstr(RS(8)&"") = "0") THEN

                       scomK = "0" & chr(9)
                       scomD = "NO&nbsp;ACTIVO" & chr(9)

                   END IF


                       scomK = scomK & "1" & chr(9) & "2"
                       scomD = scomD & "ACTIVO" & chr(9) & "CERRADO"


                       if bModificar then
                          sInput = ofv.gencomboFX(sForm,rs(6).name,rs(6),scomK,scomD," ")

                       else
                          if rs(6) = 0 then sInput = "NO&nbsp;ACTIVO"
                          if rs(6) = 1 then sInput = "ACTIVO"
                          if rs(6) = 2 then sInput = "CERRADO"
                       end if



                       celdas = celdas & ofv.GenCelda("",sc,"CENTER","20%","","","",sInput,"si")


              CASE ELSE


                   sInput = "CERRADO"
                   celdas = celdas & ofv.GenCelda("",sc,"CENTER","","","","",sInput,"si")


        END SELECT 


       sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";'  href='javascript:void(0);' "
       sinput2 = sinput2  & " onclick=" & chr(34) & "AbrirCons('PRycuentasdetTSK.asp?CTA=" & rs(0) & "'); return false;" & chr(34)
       sinput2 = sinput2  & " TITLE='DETALLE DE " & rs(2) & "'  ><b>" & sImagenC & "</b></a>" 

       celdas = celdas & ofv.GenCelda("",sc,"CENTER","","","","",sInput2,"si")



    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=PROYCUENTASDET"
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
    sAccion = "PRYCUENTASDET.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  Buscar = "" 
 
   sAccion = "PRYCUENTAS.asp"
  Celda = ofv.BotonVolver(sAccion)

 
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML

FUNCTION GENESTADO()

  XCODORDEN = 0


  StrSql = "Select MAX(ORDEN) From PROYcuentasDET WHERE CODCUENTA = 0" & XTIPO
  set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsX.eof then XCODORDEN = RSX(0)
  call ofv.cerrarconsulta(rsX)

  XCODORDEN = XCODORDEN + 10


    strsql = "insert into PROYCUENTASDET (CODCUENTA,DESCRIPCION,ORDEN,FINICIO,FFINAL,ACTIVA,TPLAN,TCONS) "
    strsql = strsql & "values ('" & XTIPO & "','NUEVA SUBCUENTA','" & XCODORDEN & "','0','99999','0','0','0')"
    
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION



%>