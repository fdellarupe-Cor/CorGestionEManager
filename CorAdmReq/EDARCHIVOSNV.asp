<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"


  sHTML = ofv.FormHeader(session("FormN"))


  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"


  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 sopcion = request.querystring("opcion")

 IF SOPCION = "SI" THEN

    GENARCHIVO()
 END IF 

  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 
  end if


    sInput = "REQUERIMIENTO: " & xCLIENTE
    celdas = ofv.GenCelda("","VtI style='background-color:#AAAAAA;color:#ffffff;'","","","","","",sInput,"si")


    filas = ofv.GenRow("","","","","25","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:2;border-color:#990000;' ","","98%","","0","","1","1","","",filas,"si") & "<BR>"


  sPath = Request.ServerVariables("PATH_INFO")
  Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1))


  Xfer   =  Xruta
  rutdir = mid(Xfer,1,instrrev(Xfer,"/",-1)) & "legajos/" & session("NCliente")
  nomdir = "Documentos-archivos-imagenes"



  rutdir = "legajos\" & session("NCliente")
  rutdir2 = "legajos/" & session("NCliente") & "/"



  'Encabezados

  formulario =  ""



  celdas = ofv.GenCelda("","VtI style='background-color:#AAAAAA;color:#ffffff;'","","","","","","ORDEN","si")
  celdas = celdas & ofv.GenCelda("","VtI style='background-color:#AAAAAA;color:#ffffff;'","","","","","","DESCRIPCION","si")
  celdas = celdas & ofv.GenCelda("","VtI style='background-color:#AAAAAA;color:#ffffff;'","","","","","","&nbsp;","si")
  formulario = ofv.GenRow("","","","","22","","",celdas,"si")

  StrSql = "Select * From reqantec where codigo = '" & session("NCliente") & "' order by orden"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "carpeta2" & rs(0)

    if rs(3) <> "NO" then
       sc = ""
    else
       sc = "uT"
    end if


    celdas = ""

       if bModificar then
          sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sform,"5","dtn","")
       else
          sInput = ofv.GenerarInput("",rs(4),"readonly","","5","dtn","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


       if bModificar then
          sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sform,"80","dt","")
       else
          sInput = ofv.GenerarInput("",rs(2),"readonly","","80","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  TARGET='_blank' href=" & chr(34) & rutdir2 & rs(3) & chr(34) &  " "
     sinput = sinput  & "  TITLE='VER'  ><b>" & sImagenI2 & "</b></a>" 

     sinputk = "&nbsp;" & sinput  & "&nbsp;"

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=ReqArcSel.asp?itemdata=" & rs(0) & "&CantRegMover=" & CantRegMoverx
     sinput = sinput  & "  TITLE='ASOCIAR'  ><b>" & sImagenI3 & "</b></a>" 

     sinputk = sinputk & sinput


    celdas = celdas & ofv.GenCelda("","uTs style='border-style:solid;border-width:1;border-color:#dddddd;' ","","","","","",sInputK,"si")


    filas = ofv.GenRow("","","","","20","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=REQANTEC"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
    rs.Movenext
  loop

    sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:2;border-color:#990000;' ","","98%","","0","","1","1","","",formulario,"si")




  'Botonera
  Agregar = ""


  cnbot = "<input type=submit class=bt value='Subir&nbsp;Documentos' >"
  Cons = ofv.gencelda("","","center","","1%","","",cnbot,"si")
  saccion = "uploadn.asp?direct="
  saccion = saccion &  rutdir & "&nomdir=" & nomdir

 ' Agregar = ofv.genform("cons","",sAccion,"_blank","post","",Cons,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
     saccion = "uploadn.asp?direct="
     saccion = saccion &  rutdir & "&nomdir=" & nomdir
     Agregar = ofv.BotonAgregar(sAccion)
  end if

'  cnbot = "<input type=submit class=bt value='Agregar&nbsp;Archivo' >"
'  Cons = ofv.gencelda("","","center","","1%","","",cnbot,"si")
'  sAccion = "EDARCHIVOSnv.asp?OPCION=SI"
'  Buscar = ofv.genform("cons","",sAccion,"","post","",Cons,"si")

  xxxx    = CorGenBotonXp(3, "NO", "NO", "NO", "NO", "Crear archivo adjunto", "NO", "CREAR")
  sAccion = "EDARCHIVOSnv.asp?OPCION=SI"


  Buscar   = ofv.GenBoton(3, "esp1Bot", "BT", "S", "", "S", "S", "N", "esp1Frm", sAccion & "")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sAccion = "EDREQUERNV.asp"
  VOLVER = ofv.BotonVolver(sAccion)


''''''''''''''''''''''

'  sHTML   = sHTML & CorPnlCOT()

'  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clpanel();' ", "S", "A", "N", "indFrm", "")

 ' Agregar = Celt1
 ' Buscar  = ""
  Celda   = Celt1

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","",Agregar,Buscar,celda,VOLVER)

''''''''''''''''''''''


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

function GENARCHIVO()



  set cnx = ofv.conectar(ofv.strconn6)


  StrSql = "Select MAX(ORDEN) From REQANTEC where codigo = '" & session("NCliente") & "' "
  set rs1 = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
  if rs1.eof then
     XORDEN = 10
  ELSE

     XORDEN = RS1(0)

     IF XORDEN <> "" THEN
        XORDEN = XORDEN + 10
     ELSE
        XORDEN = 10
     END IF

  END IF

    strsql = "insert into REQANTEC (codigo,descri,ARCHIVO,ORDEN) values ('" & session("NCliente") & "','NUEVO ARCHIVO','NO','" & XORDEN & "')"
    call ofv.crearconsultaEx(StrSql,cnx,1,parametros)



  call ofv.cerrarconn(cnx)


end function

%>