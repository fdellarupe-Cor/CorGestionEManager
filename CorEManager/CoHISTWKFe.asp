<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0




  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)
  set cn1 = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  SESSION("SEGCLIENTE") = "NO"


 DSD = request.querystring("DSD")
 sopcion = request.querystring("opcion")

  if sopcion <> "" then
     session("sopc") = sopcion
  else
     sopcion = session("sopc")
  end if 

  if DSD <> "REQ" then
     sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))
 else
     sHTML = ofv.MENUHeader("","")
 end if



'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

 
  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "window.resizeTo(790,570); "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

    celdas = ""
    filas  = ""


    sNombre = esqnomp (session("TCliente"),session("NCliente"),ofv)

  'Encabezados

     filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

     celdas =  ofv.GenCelda("","ttt colspan=4  ","","","","","",sNombre,"si")

     filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

     celdas =  ofv.GenCelda("","ut colspan=4  ","","","","","","&nbsp;","si")

     filas = filas & ofv.GenRow("","","","","","","",celdas,"si")



     celdas =  ofv.GenCelda("","tt colspan=4  ","center","","","","","HISTORIAL DE CIRCUITOS DEL DOCUMENTO","si")

     filas = filas & ofv.GenRow("","","","","25","","",celdas,"si")

     celdas =  ofv.GenCelda("","ut colspan=4  ","","","","","","&nbsp;","si")


     filas1 = filas
     filas = ""

     StrSql = "Select * From SEGCIRCUITO WHERE OBJTIPO = '" & session("TCliente") & "' and  OBJCOD = '" & session("NCliente") & "' ORDER BY ID DESC "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     posi = 0

    filas = GENLISTA(RS)

  else

     celdas = "" 

      sInput = "NO EXISTEN CIRCUITOS EN EL HISTORIAL"
      celdas = celdas & ofv.GenCelda("","vt" ,"CENTER","25%","","","",sInput,"si")
      filas = ofv.GenRow("","","","","25","","",celdas,"si")

  end if





  sHTML = sHTML  & ofv.GenTabla("","aaa  ","","98%","","0","","0","0","","",filas1 & filas,"si")


'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------



  'Botonera
  Agregar = ""

  Buscar = ""
  
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENLISTA(RSX)

    sImagenO = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=corflaticon COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=corflaticon COLOR=#008000 SIZE=3>R</FONT></B>"

    sImagenp1 = "<B><FONT FACE=corflaticon COLOR=#AA0000 SIZE=3>4</FONT></B>"
    sImagenp2 = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=3>4</FONT></B>"

  posi = 0
  prow = "N" 
  prin = sImagenp1
  Do While (Not rsX.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rsX(0)



       sr   = ""
       sc = "ut"

    xlink = "<A class=nt href='SEGCKTESTADO.asp?CKT=" & RSx(6) & "&DSD=COR"
    xlink = xlink & "' "
    xlink = xlink & " >" & prin & " " & RSx(0) & "</A>"

    sInput = xlink

    prin = sImagenp2

       celdas = ofv.GenCelda("","ut rowspan=2 style='border-style:solid;border-width:1;border-color:#f7aa22;' ","","1%","","","",sInput,"si")

       sInput2 = rsx(5)


        

       IF RSX(2) = "NO" THEN
          sInput3 = "&nbsp;PENDIENTE"
       ELSE
          sInput3 = "&nbsp;" & rsx(2)
       END IF

          sInput3 = "&nbsp;INICIO:" & sInput3

    xlink = "<A class=nt  href='SEGCKTESTADO.asp?CKT=" & RSx(6) & "&DSD=COR"
    xlink = xlink & "' "
    xlink = xlink & " >" & sInput3 & "</A>"

      celdas = celdas & ofv.GenCelda("",sc & "  style='border-right-style:solid;border-right-width:2;border-right-color:#6699cc;' ","","30%","","","",xlink,"si")


      sInput = "<b>" & sImagenC & "</b> " & rsx(8)
      celdas = celdas & ofv.GenCelda("","ut","","3%","","","",sInput,"si")


    filas = ofv.GenRow("","aaa " & sr,"","","","","",celdas,"si")

       TFIN = RSX(3) & " "

       CTFIN = INSTR(UCASE(TFIN),"CANCEL")

       IF RSX(3) = "NO" THEN
          sInput3 = "&nbsp;PENDIENTE"
       ELSEIF CTFIN > 0 THEN
          sInput3 = "&nbsp;<FONT COLOR=#993333>" & rsx(3) & "&nbsp;&nbsp;<FONT FACE=corflaticon SIZE=3>3</FONT></FONT>"
       ELSE
          sInput3 = "&nbsp;" & rsx(3)
       END IF

          sInput3 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;FIN:" & sInput3


    xlink = "<A class=nt  href='SEGCKTESTADO.asp?CKT=" & RSx(6) & "&DSD=COR"
    xlink = xlink & "' "
    xlink = xlink & " >" & sInput3 & "</A>"

      celdas = ofv.GenCelda("",sc & " style='border-right-style:solid;border-right-width:2;border-right-color:#6699cc;' ","","30%","","","",xlink,"si")


  'StrSql = "Select * From USUARIOS WHERE ID = '" & RSx(7) & "' "
  'set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  'if not rs2.eof THEN
  '     sInput = "<b>" & sImagenI & "</b> " & rs2(1)
  'END IF
     xxxx = getInfoUser(RSx(7))
      if session("SGU_nombre")<> "" then    sInput = "<b>" & sImagenI & "</b> " & session("SGU_nombre")


    celdas = celdas & ofv.GenCelda("",sc,"","50%","","","",sInput,"si")
    filas = filas & ofv.GenRow("","aaa  " & sr,"","","","","",celdas,"si")

    filas = ofv.GenTabla("","aaa  ","","100%","","0","","-2","-1","","",filas,"si")

    celdas = ofv.GenCelda("","aaa colspan=4 style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;'  ","","","","","",filas,"si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")

    formu = formu & ofv.GenForm("","","","","post","",filas,"si")


    posi = posi + 1
    rsX.Movenext
  loop

  if posi > 0 then
     celdas = "&nbsp;" 

      sInput = "<b>" & sImagenC & "&nbsp;Circuito</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","25%","","","",sInput,"si")
      sInput = "<b>" & sImagenI & "&nbsp;Usuario Iniciante</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","25%","","","",sInput,"si")

  else

     celdas = "" 

      sInput = "NO EXISTEN CIRCUITOS EN EL HISTORIAL"
      celdas = celdas & ofv.GenCelda("","vt" ,"CENTER","25%","","","",sInput,"si")


  end if

    formu = formu & ofv.GenRow("","aaa height=30","","","","","",celdas,"si")

  GENLISTA = FORMU

END FUNCTION

function EsqnomP(xtipo,xcodigo,ofv)

 set xcn=ofv.conectar(ofv.strconn2)


 IF xtipo = "RSK" THEN

    strsql = "SELECT * FROM  ORRSKCIRCUITO "
    strsql = strsql & " Where OBJID = 0" & xcodigo
    set xrs = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
    if not xrs.eof then
       IF XRS("APLIC") = "OR" THEN
          xdescri = "CIRCUITO AUTOEVALUACION"
       ELSE
          xdescri = "PLAN DE CONTINGENCIAS"
       END IF 

       Strsql = "Select * From objetos where codobjeto = 0" & XRS("CODOBJETO")
       set xrs = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
       if not xrs.eof then
          xiden = xrs(1)
          xobjdes = xrs(3)
       end if


    END IF
    call ofv.cerrarconsulta(xrs)
    
 ELSE




    Strsql="Select descri From tipobjeto where tipobjeto ='" & xtipo & "' "
    set xrs=ofv.crearconsultaEx(StrSql,xcn,1,parametros)
    if not xrs.eof then xdescri = xrs(0)
    call ofv.cerrarconsulta(xrs)

    Strsql="Select * From objetos where codobjeto =" & xcodigo
    set xrs=ofv.crearconsultaEx(StrSql,xcn,1,parametros)
    if not xrs.eof then
       xiden = xrs(1)
       xobjdes = xrs(3)
    end if
    call ofv.cerrarconsulta(xrs)

 END IF





 call ofv.cerrarconn(xcn)

 esqnomP = xdescri & ":&nbsp;" & xiden & "&nbsp;-&nbsp;" & xobjdes & ""

end function





%>