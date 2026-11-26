<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.formHeader(" ")

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

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



  sopcion = request.querystring("opcion")

  if sopcion <> "" then
     session("sopc") = sopcion
  else
     sopcion = session("sopc")
  end if 


    celdas = ""



  'Encabezados

     celdas =  ofv.GenCelda("","vti colspan=3 style='border-style:solid;border-width:1;border-color:#000000;color:#ffffff;background-color:#888888;' ","","","","","","HISTORIAL DE CIRCUITOS DEL REQUERIMIENTO","si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

 

    filas1 = filas


     StrSql = "Select * From SEGCIRCUITO WHERE  OBJCOD = '" & session("NCliente") & "' ORDER BY COMIENZO DESC "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  posi = 0

    filas = GENLISTA(RS)

  END IF








  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","1","2","","",filas1 & filas,"si")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


''''''''''''''''''''''

 ' sHTML   = sHTML & CorPnlCOT()

 ' Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clpanel();' ", "S", "A", "N", "indFrm", "")

  Agregar = Celt1
  Buscar  = ""
  Celda   = ""

  sAccion = "EDREQUERNV.asp"
  Celda = ofv.BotonVolver(sAccion)

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

''''''''''''''''''''''


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENLISTA(RSX)

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"


  posi = 0
  prow = "N" 
  Do While (Not rsX.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rsX(0)

    if prow = "S" then
       prow = "N"
       sr   = " bgcolor=#eeeeee "
    else 
       prow = "S"
       sr   = ""
    end if

       sr   = ""
       sc = "vt"

      celdas = ""

    xlink = "<A class=cc href='SEGCKTESTADONV.asp?CKT=" & RSx(6) & "&DSD=REQ"
    xlink = xlink & "' "
    xlink = xlink & " ><IMG  src='Imagenes/ICorAbierto.jpg'  border=0 ></A>"

    sInput = xlink



       sInput2 = rsx(5)

       StrSql = "Select * From REQUERIMIENTO where codigo = '" & rsx(5) & "' "
       set rs2 = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
       if NOT rs2.eof then
          sInput2 = "<B>xxxx" & RS2(2) & "</B>"
       end if

       IF RSX(3) = "NO" THEN
          sInput3 = "&#0183;&nbsp;<font color='#990000' >PENDIENTE</font>"
       ELSE
          sInput3 = "&#0183;&nbsp;" & rsx(3)
       END IF

       sInput4 = rsx(8)
       sInput5 = "&#0183;&nbsp;" & rsx(2)


    xlink = "<A class=cc  href='SEGCKTESTADONV.asp?CKT=" & RSx(6) & "&DSD=REQ"
    xlink = xlink & "' "
    xlink = xlink & " >&#0183;&nbsp;" & sInput4 & "</A>"







   ''   celdas = celdas & ofv.GenCelda("","ut rowspan=2 ","","1%","","","",sInput,"si")


      celdas = celdas & ofv.GenCelda("","uti rowspan=2 style='border-style:solid;border-width:1;border-color:#888888;' ","","50%","","","",xlink,"si")

      celdas = celdas & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:#888888;background-color:#f7eedd;' ","","25%","","","",sInput5,"si")

      celdas = celdas & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:#888888;background-color:#f7eedd;' ","","25%","","","",sInput3,"si")



    filas = ofv.GenRow("","aaa ","","","20","","",celdas,"si")

  StrSql = "Select * From USUARIOS WHERE ID = '" & RSx(7) & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if not rs2.eof THEN

       sInput = "&#0183;&nbsp;" & rs2(1)
  END IF
       celdas = ofv.GenCelda("","uti colspan=2 style='border-style:solid;border-width:1;border-color:#888888;' ","","50%","","","",sInput,"si")


    filas = filas & ofv.GenRow("","aaa  ","","","20","","",celdas,"si")

    celdas = ofv.GenCelda("","uti colspan=3 style='border-top-style:solid;border-top-width:1;border-top-color:#bbbbbb;' ","","50%","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","aaa  ","","","10","","",celdas,"si")

 '   filas = ofv.GenTabla("","","","100%","","0","","-2","-1","","",filas,"si")

'    celdas = ofv.GenCelda("","aaa colspan=4 style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;'  ","","","","","",filas,"si")
 

 '''  filas = ofv.GenRow("","","","","","","",celdas,"si")

    formu = formu & ofv.GenForm("","","","","post","",filas,"si")


    posi = posi + 1
    rsX.Movenext
  loop


  GENLISTA = FORMU

END FUNCTION




%>