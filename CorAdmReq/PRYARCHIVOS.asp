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


  sHTML = replace(ofv.FormHeader(session("FormN")),"<body ","<body onload='uno();' ")  & chr(13)
  shtml = shtml & "<script  language=javascript>" & chr(13)
  shtml = shtml & " var cnt = 0; " & chr(13)
  shtml = shtml & " function uno() {  " & chr(13)
  shtml = shtml & " if (window.name=='info') "
  shtml = shtml & " { document.all.DOS.style.display='block';} else  { document.all.DOS.style.display='none';}  "
  shtml = shtml & " if (document.all.odiv) { document.all.odiv.style.top=document.body.clientHeight - document.all.odiv.clientHeight; " & chr(13)
  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
  shtml = shtml & " setTimeout (" & chr(34) & "uno()" & chr(34) & ", 350); }}" & chr(13)
  shtml = shtml & " } " & chr(13)
  shtml = shtml & " </script>" & chr(13)

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

  StrSql = "Select * From PROYECTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 
  end if

    sInput = "PROYECTO:"
    celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

    sInput = session("NCliente") & " - " & xCLIENTE
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"





  'Encabezados
''  formulario =  ofv.GenCelda("","VT colspan=5","center","","","","","Solapas","si")

  formulario =  ""
''  formulario =  formulario & ofv.Encabezados("Orden","Descripcion","&nbsp;","&nbsp;")

  celdas = ofv.GenCelda("","vt  ","","","","","","ORDEN","si")
  celdas = celdas & ofv.GenCelda("","vt  ","","","","","","DESCRIPCION","si")
  celdas = celdas & ofv.GenCelda("","vt  ","","","","","","&nbsp;","si")
  formulario = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas =  ofv.GenCelda("","VT colspan=3 style='border-TOP-style:ridge;border-TOP-width:2;border-TOP-color:#f7aa22;' ","center","","","","","&nbsp;","si")
  formulario =  formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  StrSql = "Select * From PROYantec where codigo = '" & session("NCliente") & "' order by orden"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "carpeta2" & rs(0)

    if rs(3) <> "NO" then
       sc = ""
    else
       sc = "VT"
    end if

 '''   sInput1 = ofv.GenerarInput(rs(0).name,rs(0),"hidden",sform,"10","dt","")
 ''   celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")

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
     sinput = sinput  & "  TARGET='_blank' href='PRYImagenCont.asp?Solapa=" & rs(3) & "' "
     sinput = sinput  & "  TITLE='VER'  ><b>" & sImagenI2 & "</b></a>" 

     sinputk = sinput  & "&nbsp;"

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PRYArcSel.asp?itemdata=" & rs(0) & "&CantRegMover=" & CantRegMoverx
     sinput = sinput  & "  TITLE='ASOCIAR'  ><b>" & sImagenI3 & "</b></a>" 

     sinputk = sinputk & sinput  & "&nbsp;"


    celdas = celdas & ofv.GenCelda("","VT","","","","","",sInputK,"si")



 '   sinput = "<a TARGET='_blank' href='ImagenCont.asp?Solapa=" & rs(3) & "' "
 '   sinput = sinput  & " ><b>VER</b></a>" 
 '   celdas = celdas & ofv.GenCelda("","TT","","","","","",sInput,"si")

 '   sinput = "<a  href=ReqArcSel.asp?itemdata=" & rs(0) & "&CantRegMover=" & CantRegMoverx
 '   sinput = sinput  & " ><b>ASOCIAR</b></a>" 
 '   celdas = celdas & ofv.GenCelda("","TT","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=PROYANTEC"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
    rs.Movenext
  loop

 '' sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")
    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=#dddddd,direction=125);' ","","98%","","0","","0","0","","",formulario,"si")


  sPath = Request.ServerVariables("PATH_INFO")
  Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1))


  Xfer   =  Xruta
  rutdir = mid(Xfer,1,instrrev(Xfer,"/",-1)) & "PROYECTOS/" & session("NCliente")
  nomdir = "Documentos-archivos-imagenes"


  'Botonera

  cnbot = "<input type=submit class=bt value='VOLVER' >"
  Cons = ofv.gencelda("DOS","","center","","1%","","",cnbot,"si")
  sAccion = "PROYPROP.asp"
  VUELVE = ofv.genform("","",sAccion,"","post","",Cons,"si")


  Agregar = ""
  if bAgregar then
     cnbot = "<input type=submit class=bt value='Subir&nbsp;Documentos' >"
     Cons = ofv.gencelda("","","center","","1%","","",cnbot,"si")
     saccion = "uploadn.asp?direct="
     saccion = saccion &  rutdir & "&nomdir=" & nomdir
     Agregar = ofv.genform("cons","",sAccion,"_blank","post","",Cons,"si")
  END IF

  Buscar = ""
  if bAgregar then
     cnbot = "<input type=submit class=bt value='Agregar&nbsp;Archivo' >"
     Cons = ofv.gencelda("","","center","","1%","","",cnbot,"si")
     sAccion = "PRYARCHIVOS.asp?OPCION=SI"
     Buscar = ofv.genform("cons","",sAccion,"","post","",Cons,"si")
  END IF

  fill = space(50)

  fill = replace(fill," ","&nbsp;")

  fill = ofv.gencelda("","vt","right","","100%","","",fill,"si")


  sHTML = sHTML & "<div  id=odiv name=odiv style='position:absolute;top:-10px;filter:shadow(color=" & Session("APLSHW") & ",direction=125);'  >"
  sHTML = sHTML & replace(ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",VUELVE,Agregar,Buscar,fill),"id=btnas width=2%","id=btnas width=100% style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ")
  sHTML = sHTML & "</div>"

  
 '' sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

function GENARCHIVO()



  set cnx = ofv.conectar(ofv.strconn6)


  StrSql = "Select MAX(ORDEN) From PROYANTEC where codigo = '" & session("NCliente") & "' "
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

    strsql = "insert into PROYANTEC (codigo,descri,ARCHIVO,ORDEN) values ('" & session("NCliente") & "','NUEVO ARCHIVO','NO','" & XORDEN & "')"
    call ofv.crearconsultaEx(StrSql,cnx,1,parametros)



  call ofv.cerrarconn(cnx)


end function

%>