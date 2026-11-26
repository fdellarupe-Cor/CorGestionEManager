<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
   'response.write CorStyler2()
  'sHTML = ofv.FormHeader(session("FormN"))
  response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
   response.write "<tr>"
   response.write "<td width=100% >"
  response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
   response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
   response.write "<tr>"
   response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
   response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
   response.write "</tr></table>"


  'Encabezados
  'Formulario = ofv.Encabezados("Codigo","Descripcion")
  Formulario=chr(13)&" <TR ><TD  class=tt width=2% >Perfil</TD><TD  class=tt width=98% >Descripcion</TD></tr>" & chr(13)


  'ofv.ObtenerAtributos Session("Form"),""
  'bAgregar = ofv.AAgregar
  'bModificar = ofv.AModificar
  'bSuprimir = ofv.AEliminar
  'bConsultar = ofv.AConsultar

  'call ofv.cerrarconn(cn)

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From TipoPerfil Order By descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  'if Not rs.EOF then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) 'and (posi < CantRegMostrarx)
  '  sForm = "form" & posi


      Celda = ""

    'Perfil
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","vt","","50","","","",sInput,"si")

    'Descripcion
    sInput = rs(1) 'ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"125","dt","")
    Celda = Celda & ofv.gencelda("","vt","","","","","","<u>" & sInput & "</u>","si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")
    Formulario = Formulario & Fila

      Strsql = "SELECT ECO.ECO ,ECO.descri "
      Strsql = Strsql & "FROM  ECO, ECOPERFIL, TIPOPERFIL "
      Strsql = Strsql & " where  "
      Strsql = Strsql & " ECO.eco = ECOPERFIL.codsector and "
      Strsql = Strsql & " ECOPERFIL.codperfil = TIPOPERFIL.tipoperfil "
      Strsql = Strsql & " and ECOPERFIL.codperfil = '" & rs(0) & "' order by ECO.descri "
      set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if Not rs1.EOF then
         Do While (Not rs1.EOF)

            Celda = ""

           'Perfil
            sInput = " "
            Celda = Celda & ofv.gencelda("","vt","right","50","","","",sInput,"si")

           'Descripcion
            sInput = "&nbsp;*&nbsp;" & rs1(1)
            Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")

            Fila = ofv.genrow("","","","","","","",Celda,"si")
            Formulario = Formulario & Fila

         rs1.movenext
         loop
            sInput = "&nbsp;"
            Celda = ofv.gencelda("","ut colspan=2","","","","","",sInput,"si")
            Fila = ofv.genrow("","","","","","","",Celda,"si")
            Formulario = Formulario & Fila
      end if

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  'Botonera
 ' Agregar = ""
 ' if bAgregar then
 '   sAccion = "AgregarTipoPerfil.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=TipoPerfil"
 '   Agregar = ofv.BotonAgregar(sAccion)
 ' end if

 ' str1 = "Select descri From Tipoperfil Order By descri "    ' orden pagina
 ' str2 = "Select descri From Tipoperfil Order By descri "    ' orden alfabetico
 ' Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
 ' 'consulta
 ' sinput = "<a  href=PerfilesdeComponentesCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
 ' Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
 '
  'sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = shtml & "</body></html>"
end if

Response.Write sHTML

%>