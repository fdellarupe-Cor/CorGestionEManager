<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = request.querystring("CantRegMostrar")
  CantRegMoverx = request.querystring("CantRegMover")


  wperfil = request.querystring("wperfil")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_wperfil=" & wperfil

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ""


  dim stabla(12)


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  posi = 0

  Strsql = "Select * From Objperfil where tipoperfil = '" & wperfil & "' "
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     if clng(rs(32)) = 9999 then

  Strsql = "UPDATE OBJPERFIL SET objdsc01 = 'PROPIEDADES',objdsc02 = 'CONSIDERACIONES GENERALES',objdsc03 = 'SECTORES',"
  Strsql = strsql & "objdsc04 = 'EVENTOS',objdsc05 = 'PROCEDIMIENTOS',objdsc06 = 'DESCRIPCION',objdsc07 = 'CONTROLES',"
  Strsql = strsql & "objdsc08 = 'ESQUEMAS',objdsc09 = 'TABLAS',objdsc10 = 'DOCUMENTOS',objdsc11 = 'CONTENIDO',objdsc12 = 'CURSOGRAMAS',"
  Strsql = strsql & "objdsc13 = 'CONTRATOS',objdsc14 = 'PREGUNTAS FRECUENTES',objdsc15 = 'GENERAL',objord01 = 2,objord02 = 5,objord03 = 4,objord04 = 6,"
  Strsql = strsql & "objord05 = 7,objord06 = 3,objord07 = 8,objord08 = 9,objord09 = 12,objord10 = 11,objord11 = 10,"
  Strsql = strsql & "objord12 = 13,objord13 = 14,objord14 = 15,objord15 = 1"
  Strsql = strsql & " where tipoperfil = '" & wperfil & "' "

  call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     end if

  end if

  call ofv.cerrarconsulta(rs)


  Strsql = "Select * From Objperfil where tipoperfil = '" & wperfil & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     fila = ""


      Celda = ofv.GenCelda("","ttt","","15%","","","","Perfil: ","si")
      sInput = rs(0) & "&nbsp;-&nbsp;" & rs(1)

      Celda = celda & ofv.GenCelda("","vti","","90%","","","",sInput,"si")

      sinput = "<a  href=PerfilesdeEsquemaNormativoOrd.asp?CantRegMover=" & CantRegMoverx
      sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & "&Wperfil=" & Wperfil  & "&Welem=0 ><b>Orden</b></a>"
  '    Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

      sinput = "<a  href=PerfilesdeEsquemaNormativo.asp?CantRegMover=" & CantRegMoverx
      sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & " ><b>Volver</b></a>"
  '    Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")



      Formulario = Fila

      sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","80% ","","0","","0","1","","",Formulario,"si") & ""



      fila = ""

      Celda = ofv.GenCelda("","tt ","","","","","","solapa","si")
      Celda = celda & ofv.GenCelda("","tt","","","","","","titulo","si")
      Celda = celda & ofv.GenCelda("","tt","","","","","","activa","si")
      Celda = celda & ofv.GenCelda("","tt","","","","","","descripcion","si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
      Formulario = Fila

      Formulario = ofv.Encabezados("solapa","titulo","activa","descripcion")

dim atab(15)

    for xx = 1 to 15
           atab(xx) = 99
    next

    for xx = 2 to 16

        if clng(rs(xx + 30)) < 99 then

           atab(rs(xx + 30)) = xx

        end if


    next

    for yy = 1 to 15

        if clng(atab(yy)) = 99 then exit for

        ''response.write atab(yy) & " atab(yy) " & "<br>"

           xx = atab(yy)
           fila = ""
           sForm = "form" & posi


           Celda = ""
           sInput = rs(xx + 30)
           Celda = celda & ofv.GenCelda("","CC colspan=2 ","right","1%","","","",sInput,"si")

           if rs(xx)  then
              clase = "ut style='background-color:color-mix(in srgb, #00bb00, transparent 60%);' "
           else
              clase = "ut"
           end if


           sInput = ofv.GenerarInput(rs(xx + 15).name,rs(xx + 15),"text",sForm,"60","dt","")

        '' response.write rs(xx + 15).name & " name " & "<br>"
        '' response.write rs(xx + 15)      & " indice " & "<br>"

           Celda = celda & ofv.GenCelda("",clase,"","40%","","","",sInput,"si")

           if yy = 1 then
              sInput = "<b>SI</b>"
           else
              sInput = ofv.combosino(sForm,rs(xx),rs(xx).name)
           end if




           Celda = celda & ofv.GenCelda("",clase,"center","5%","","","",sInput,"si")



           sInput = "&nbsp;"

           Strsql = "Select descri From Objperfildsc where dato = '" & rs(xx).name & "' "
           set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rs2.eof then
              sInput = rs2(0)
           end if
           call ofv.cerrarconsulta(rs2)
		   

           Celda = celda & ofv.GenCelda("",clase,"","40%","","","","<i>" & sInput & "</i>","si")

           Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

           sAccion = ""
           if bModificar then
              sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
              sAccion = sAccion & "&tabla=Objperfil&volver=" & vuelta
           end if
           Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

           posi = posi + 15



    next

  end if
  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 0px 0px 10px 10px;'","","80%","","0","","0","0","","",filas,"si")




  Agregar = ""
  Buscar  = ""

  sAccion = "PerfilesdeEsquemaNormativo.asp?CantRegMover=" & CantRegMoverx
  Celda   = ofv.BotonVolver(sAccion)

  xxxx    = CorGenBotonXp(16, "NO", "NO", "NO", "NO", "Orden", "_self", "Orden")
  sAccion = "PerfilesdeEsquemaNormativoOrd.asp?CantRegMover=" & CantRegMoverx & "&Wperfil=" & Wperfil  & "&Welem=0"
  Listar  = ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar & Listar,celda)





  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>