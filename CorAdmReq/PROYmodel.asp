<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  sopcion = request.querystring("opcion")

    if sopcion = "SI" then
       XXXX = GENPROY("NEW")

    end if


 LISTA = ""
 DESTINO = ""

  STIPO  = "PRY" 
    STABLA = "PROYECTO"
    STABLA2 = "PROYRESP"

    LISTA = "  MODELOS DE PROYECTOS "

    XDEST="GLOBAL"

  Strsql = "Select * from clienteplan where id = 1"
  set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
  if not rs.eof then
     NOMEMP = rs(1)
     LOGEMP = rs(5)
  end if
  call ofv.cerrarconsulta(rs)

       XDESTN=NOMEMP


     StrSql3 = "Select * From " & STABLA & " WHERE " & STABLA & ".CODIGO <> ''  "

     STRSQL3 = STRSQL3 & "  AND " & STABLA & ".ESTADO = 99 ORDER BY  " & STABLA & ".FECHAINI desc "

     StrSql2 = "Select DISTINCT ESTADO From " & STABLA & "  ORDER BY " & STABLA & ".ESTADO "


    DESTINO = "&nbsp;"


    sImagenO  = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI  = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE  = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC  = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"

    sImagenX  = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>4</FONT></B>"
    sImagenIX = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>2</FONT></B>"



 DIM ATABLA(50,2) 

 ITAB=0
 TTAB=1 
 ITABB=0


    celdas = ""


  'Encabezados
  formulario = formulario & ofv.GenCelda("","tt colspan=5","center","","","","",LISTA & DESTINO,"si")


  IF TTAB = 0 THEN

        celdas = ofv.GenCelda("","VT colspan=5","center","","","","","NO HAY " & STABLA & "S PARA LA SELECCION ","si")
        formulario = formulario & ofv.GenRow("","aaa  ","","","","","",celdas,"si")

  ELSE



     STRSQL = STRSQL3 
 ''    response.write strsql & "<br>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        rs.move(CantRegMoverx)
  posi = 0
        formulario = formulario & GENLISTA(RS)

      celdas = ""



     ELSE

        celdas = ofv.GenCelda("","VT colspan=5","center","","","","","NO HAY " & STABLA & "S PARA LA SELECCION ","si")
        formulario = formulario & ofv.GenRow("","aaa  ","","","","","",celdas,"si")

     END IF


  END IF


  sHTML = sHTML  & ofv.GenTabla("","aaa style='filter:shadow(color=#cccccc,direction=125);' ","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "PROYMODEL.asp?Opcion=SI&OPC=1&TIPO=" & STIPO & "&RESP=" & SRESP & "'  "
    Agregar = REPLACE(ofv.BotonAgregar(sAccion),"Agregar","NUEVO MODELO")
  end if

  Buscar = ""
  

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML








FUNCTION GENLISTA(RSX)




      formu = ""


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

    if bSuprimir  then
      sHref = "PRYREQDELETE.asp?ID=" & rsX(1) & "&CantRegMover=" & CantRegMoverx
      CeldaS = ofv.BotonEliminar( "MODELO : " & rsX(2) & " - (ID:" & RSX(1) & ")",sHref)
    else
      CeldaS = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if





       xlink = "<a class=cc   href=proyMODprop.asp?pry=" & rsx(1)
       xlink = xlink & " ><IMG  src='../CORESS20/Imagenes/ICorONTIME.jpg'  border=0 ></A>"

       xlink2 = "<a class=cc   href=proyMODprop.asp?pry=" & rsx(1)
       xlink2 = xlink2  & " ><b>" & rsX(2) & "</b></a>" 



      sInput = xlink

      celdas = celdas & ofv.GenCelda("","vt","","1%","","","",sInput,"si")



      sInput = xlink2

      celdas = celdas & ofv.GenCelda("","vt","","60%","","","",sInput,"si")


      sInput = "<b>" & sImagenC & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","1%","","","",sInput,"si")




    if clng(rsx(3)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rsx(3) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha
      celdas = celdas & ofv.GenCelda("",sc & " ","","37%","","","",sInput,"si")







    filas = ofv.GenRow("","aaa " & sr,"","","","","",celdas,"si")




    formu = formu & filas

      celdas = ofv.GenCelda("","AAA COLSPAN=5 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;FONT-SIZE:5pt' ","","1%","","","","&nbsp;","si")

      formu = formu & ofv.GenRow("","aaa " & sr,"","","","","",celdas,"si")



    posi = posi + 1
    rsX.Movenext
  loop



  GENLISTA = FORMU

END FUNCTION

Function GENPROY(YPRY)



'' Set CN = OFV.conectar(OFV.strconn6)

    If Session("PRCF") <> "S" Then
       xxxx = OWKF.GETPRYCNFG()
    End If



        CODREQ = OWKF.GTPROY()

           Yxrtit = "NUEVO MODELO"
           Yxrtipo = "0"
           Yxrorig = "0"
           Yxruori = Session("USUARIO")
           Yxruasi = Session("USUARIO")
           Yxrts = "0"
           YXTID = "NO"
           YXRID = "0"
           YXPRI = "3"

 If nFecini < CLng(CDate("01-01-1900")) Then
    nFecini = CLng(CDate("01-01-1900"))
 Else
    nFecini = 0
 End If

     vNuevoValor = CLng(Date) - nFecini
     YXFHORA = Right(Now(), Len(Now()) - InStr(Now(), " "))

     StrSql = "insert into PROYECTO (CODIGO,TITULO,FECHAINI,FECHASOL,FECHAENT,FECHAFIN,ORIGEN,TIPO,"
     StrSql = StrSql & "TEMASUBTEMA,DESCRIPCION,ESTADO,DESCTIPO,ORIGINANTE,ASIGNADO,TIPOPROY,OBJPROY,VPLAN,TFECHA,PRIORIDAD) "
     StrSql = StrSql & "values ('" & CODREQ & "','" & Yxrtit & "','" & vNuevoValor & "','0','0','0',"
     StrSql = StrSql & "'" & Yxrorig & "','" & Yxrtipo & "','" & Yxrts & "','MODELO DE PROYECTO','99','NO',"
     StrSql = StrSql & "'" & Yxruori & "','" & Yxruasi & "','" & YXTID & "','" & YXRID & "','0','" & Session("DHC") & "','" & YXPRI & "')"
     Call OFV.crearconsultaEx(StrSql, CN, 1, "")

     StrSql = "INSERT INTO PROYPLAN (CODIGO,VPLAN,DESCRIPCION,FECHAPLAN,HORAPLAN,USRRESP,ESTADO) VALUES ("
     StrSql = StrSql & "'" & CODREQ & "','0','PLAN ORIGINAL','" & vNuevoValor & "','" & YXFHORA & "','" & Yxruasi & "','1')"
     Call OFV.crearconsultaEx(StrSql, CN, 1, "")



'' Call OFV.cerrarconn(CN)

End Function

%>