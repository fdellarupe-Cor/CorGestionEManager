<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok  
else
  response.expires=0 

  
  sHTML =   ofv.formHeader("") 

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "status=yes,resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"
    
  CantRegMostrarx = 6 
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
  sopc    = request.querystring("opc")    
  sTIPO   = request.querystring("TIPO")
  sRESP   = request.querystring("RESP") 
  sSEL    = request.querystring("SEL")
  XOPT    = request.querystring("XOPT")
  XRESP   = request.querystring("RESP")
  Xestado   = request.form("estado")
  Xfdesde   = request.form("fdesde")
  Xfhasta   = request.form("fhasta")
  XORIGINANTE=request.form("XORIGINANTE")
  XANALISTA=request.form("XANALISTA")
  xgersector=request.form("xgersector")
  xBusca=  ucase(request.querystring("busca")&"")
  'response.write request.querystring
 
 
             LISTA = " LISTA DE  REQUERIMIENTOS"

 
    sImagenO  = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI  = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE  = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC  = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"

    sImagenX  = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>4</FONT></B>"
    sImagenIX = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>2</FONT></B>"

    sImagenpry = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>&#0143;</FONT></B>"
    sImagenreq = "<B><FONT FACE=WEBDINGS COLOR=#006600 SIZE=3>&#0153;</FONT></B>"
 



 DIM ATABLA(50,2) 

 ITAB=0 
 TTAB=0 
 ITABB=0



    celdas = "" 



  'Encabezados
  formulario = formulario & ofv.GenCelda("","ttT colspan=7","","","","","",LISTA & DESTINO,"si")


  select case xBusca
   case "ESTADO"
        strsql="Select * From REQUERIMIENTO WHERE ESTADO = " & Xestado
   
   case "RANGO"     
       strsql="Select * From REQUERIMIENTO WHERE FECHAINI BETWEEN " & XFDESDE & " AND " & XFHASTA
       
       
   case "GERSECTOR"     
       strsql="Select a.* From REQUERIMIENTO a ,reqsectores b " 
       strsql=strsql & " where a.codigo=b.codigo and (b.gerencia=0"& xgersector &" or b.sector=0"&xgersector&")"  
       
   case "ANALISTA"     
       strsql="Select * From REQUERIMIENTO WHERE ASIGNADO='" & XANALISTA &"'"
   
   case "ORIGINANTE"     
       strsql="Select * From REQUERIMIENTO WHERE ORIGINANTE='" & XORIGINANTE & "'"
  end select
    
    'response.write strsql & "<br>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        rs.move(CantRegMoverx)
        posi = 0
        formulario = formulario & GENLISTA(RS)

      celdas = ""

      sInput = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","vt colspan=3","","","","","",sInput,"si")

      sInput = "<b>" & sImagenC & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")


      sInput = "<b>&nbsp;Fecha Creacion</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")

      sInput = "<b>" & sImagenE & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

      if stipo = "PRY" then
         sInput = "<b>&nbsp;Fecha Inicio</b>" 
      else
         sInput = "<b>&nbsp;Fecha Solicitud</b>"
      end if
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")

      formulario = formulario & ofv.GenRow("","aaa height=30  ","","","","","",celdas,"si")

      celdas = ""

      sInput = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","vt colspan=3","","","","","",sInput,"si")

      sInput = "<b>" & sImagenO & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

      sInput = "<b>&nbsp;Usuario Originante</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")

      sInput = "<b>" & sImagenI & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

      sInput = "<b>&nbsp;Usuario Interviniente</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")

      formulario = formulario & ofv.GenRow("","aaa  ","","","","","",celdas,"si")

     ELSE

        celdas = ofv.GenCelda("","VT colspan=7","center","","","","","<BR><BR><BR>NO HAY " & STABLA & "S PARA LA SELECCION<BR><BR><BR>","si")
        formulario = formulario & ofv.GenRow("","aaa  ","","","","","",celdas,"si")

     END IF


 ' END IF


  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","-1","-1","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "PROYLISTDSP.asp?Opcion=SI&OPC=1&TIPO=" & STIPO & "&RESP=" & SRESP & "'  "
    Agregar = REPLACE(ofv.BotonAgregar(sAccion),"Agregar","NUEVO " & STABLA)
  end if

  Buscar = ""
   Agregar = ""

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML


FUNCTION GENREQ(XTIPOREQ)

    CODREQ = GTREQ()



if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini

     XSECTORIG  = "0"
     XSECTORIGN = 0
     XCODINTER  = "0"

     StrSql = "Select * From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
     if NOT rsX2.eof then
        XSECTORIGN = rsX2(21)
        XSECTORIG  = CSTR(rsX2(21))
     END IF


     IF XSECTORIGN > 0 THEN

     StrSql = "Select CODINTER From ECO where ECO = " & XSECTORIGN
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        XCODINTER = rsX2(0)
     END IF

     END IF


     IF XCODINTER = "" THEN XCODINTER = "0"


     XYEAR = datepart("yyyy",date())

      

     StrSql = "Select SECUEN From REQCODINT WHERE COD1 = '" & XCODINTER & "' AND COD2 = '" & XYEAR & "' "
   ''  response.write strsql & "<br>"
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsX2.eof then

        XSECUEN = rsX2(0) + 1

        strsql = "UPDATE REQCODINT SET SECUEN = " & XSECUEN
        strsql = strsql & " WHERE COD1 = '" & XCODINTER & "' AND COD2 = '" & XYEAR & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     ELSE

        XSECUEN = 1
        
        strsql = "insert into REQCODINT (COD1,COD2,COD3,COD4,COD5,SECUEN) "
        strsql = strsql & "values ('" & XCODINTER & "','" & XYEAR & "','NO','NO','NO','1')"
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     END IF

    FOR YYY=LEN(XSECUEN) TO 5

        XSECUEN = " " & XSECUEN


    NEXT 


    XSECUEN = REPLACE(XSECUEN," ","0")


    XCODINT = XCODINTER & "-" & XSECUEN & "-" & XYEAR


    strsql = "insert into REQUERIMIENTO (CODIGO,TITULO,FECHAINI,FECHASOL,FECHAENT,FECHAFIN,ORIGEN,TIPO,"
    strsql = strsql & "TEMASUBTEMA,DESCRIPCION,ESTADO,DESCTIPO,ORIGINANTE,ASIGNADO,TIPOREQ,FECHATEN,PRIORIDAD,CODINT) "
    strsql = strsql & "values ('" & CODREQ & "','NUEVO REQUERIMIENTO','" & vNuevoValor & "','0','0','0','0','0','0','DESCRIPCION','1','NO','" & SESSION("USUARIO") & "','NO','" & XTIPOREQ & "','0','3','" & XCODINT & "')"
   '' response.write strsql & "<br>"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)



''' crea subdirectorio de la carpeta

  Set fso = CreateObject("Scripting.FileSystemObject")

  sPath = server.mappath("./Legajos")

  Set Folder = fso.getfolder(sPath)
  set sfolder = Folder.SubFolders
  sfolder.add(CODREQ)
  xxx = GENSECTOROR(CODREQ)



END FUNCTION

FUNCTION GENSECTOROR(XCODREQ)



  StrSql = "Select * From REQUERIMIENTO where codigo = '" & XCODREQ & "' "
  set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rsX.eof then
     XTIPOREQ = RSX(15)
     XUSRORIG = RSX(13)
     XSECTORIG = 0

     StrSql = "Select * From USUARIOS where ID = '" & XUSRORIG & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
     if NOT rsX2.eof then
        XSECTORIG = rsX2(21)
        XSECTORIGB = XSECTORIG
        XSECTGEREN = 0

        DO
          'Sector Principal
          StrSql = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1 where ECO.eco = ECOREL.ECODEPEND"
          StrSql = StrSql & " and ECOREL.ECOMASTER = ECO_1.eco and ECO.eco = " & XSECTORIGB
''RESPONSE.WRITE STRSQL
          set rsX3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
          If rsX3.EOF Then
             EXIT DO
          Else
              XSECTGEREN = rsX3(0)
              XSECTORIGB = rsX3(0)
              StrSql = "SELECT * FROM CSECOSNJ where ECO = " & XSECTORIGB
              set rsX4 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
              If rsX4.EOF Then
                 EXIT DO
              Else
                 XJERAR = RSX4(3)
                 XJERAR = INSTR(XJERAR,"GEREN")
                 IF XJERAR > 0 THEN
                    EXIT DO
                 END IF
              End If
          End If
        
        LOOP 

        STRSQL = "INSERT INTO REQSECTORES (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC) VALUES ("  
        STRSQL = STRSQL & "'" & XCODREQ & "','10','" & XSECTORIG & "','" & XSECTGEREN & "','OR','','" & XUSRORIG & "')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

     END IF


     StrSql = "Select * From TIPOREQ where TIPOREQ = '" & XTIPOREQ & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        XSECTORIG  = rsX2(34)
        XSECTGEREN = rsX2(33)

        STRSQL = "INSERT INTO REQSECTORES (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC) VALUES ("  
        STRSQL = STRSQL & "'" & XCODREQ & "','20','" & XSECTORIG & "','" & XSECTGEREN & "','AN','','NO')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

     END IF







  END IF



END FUNCTION





FUNCTION GENLISTA(RSX)



       sInput2 = sImagenX & "<b>&nbsp;ESTADO:&nbsp;" & ATABLA(ITABB,2) & "</b>&nbsp;" & sImagenIx

       sinput = "<a class=cc  STYLE='CURSOR:HAND;' href'javascript:void(0);' "
       sinput = sinput  & " title='Ver&nbsp;estados' onclick=" & chr(34)
       sinput = sinput  & "document.all.mn1.style.top=event.clientY;"
       sinput = sinput  & "document.all.mn1.style.left=event.clientX;document.all.mn1.style.display='block';"
       sinput = sinput  & chr(34) & " ><b>" & sInput2 & "</b></a>" 

''alert(event.clientX + ' ' + event.clientY);


      celdas = ofv.GenCelda("xx1","vt COLSPAN=7 style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#DDDDDD;' ","","1%","","","",sInput,"si")

      formu = formu & ofv.GenRow("","","","","30","","",celdas,"si")

   xTrColor = "#FFFFFF"

  posi = 0
  prow = "N" 
  Do While (Not rsX.EOF) 'and (posi < CantRegMostrarx)

    sform = "REQ" & rsX(0)

    if prow = "S" then
       prow = "N"
       sr   = " bgcolor=#eeeeee "
    else 
       prow = "S"
       sr   = ""
    end if

       sr   = ""
       sc = "ut"

    if bSuprimir and ATABLA(ITABB,1) = 1 then
      sHref = "PRYREQDELETE.asp?ID=" & rsX(1) & "&CantRegMover=" & CantRegMoverx
      CeldaS = ofv.BotonEliminar( STABLA & ": " & rsX(2) & " - (ID:" & RSX(1) & ")",sHref)
    else
      CeldaS = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if



    IF STIPO = "PRY" THEN

       xlink = "<A  href='&nbsp;' class=nt "
     '''  xlink = xlink & " onclick=" & chr(34) & "AbrirCons('ConsPRY.asp?codigo=" & rsx(1) & "'); return false;" & chr(34)

       xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" & rsx(1) & "&tipo=PRY'); return false;" & chr(34)
       xlink = xlink & " >" & sImagenpry & "</A>"

''       xlink2 = "<a class=cc   href=proyprop.asp?pry=" & rsx(1)
''       xlink2 = xlink2  & " ><b>" & rsX(2) & "</b></a>" 

       xlink2 = "" & rsX(2) & "" 

    ELSE

       xlink = "<A  href='&nbsp;' class=nt "
'''       xlink = xlink & " onclick=" & chr(34) & "AbrirCons('Consultar.asp?codigo=" & rsx(1) & "&tipo=" & STIPO  & "'); return false;" & chr(34)
       xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" & rsx(1) & "&tipo=REQ'); return false;" & chr(34)
       xlink = xlink & " >" & sImagenreq & "</A>"

       xlink2 = "" & rsX(2) & "" 

    END IF

      sInput = xlink

      celdas = celdas & ofv.GenCelda("","ut","","1%","","","",sInput,"si")



      sInput = xlink2

      celdas = celdas & ofv.GenCelda("","ut","","30%","","","",sInput,"si")


      sInput = "<b>" & sImagenC & "</b>"
      celdas = celdas & ofv.GenCelda("","ut","","1%","","","",sInput,"si")




    if clng(rsx(3)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rsx(3) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha
      celdas = celdas & ofv.GenCelda("","ut","","37%","","","",sInput,"si")

      sInput = "<b>" & sImagenE & "</b>"
      celdas = celdas & ofv.GenCelda("","ut","","1%","","","",sInput,"si")

    if clng(rsx(4)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rsx(4) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha



      celdas = celdas & ofv.GenCelda("","ut","","37%","","","",sInput,"si")



    filas = ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")


      celdas = ofv.GenCelda("","vt COLSPAN=2 ","","1%","","","","","si")

    IF STIPO = "PRY" THEN

       sinput = "&nbsp;"

    ELSE

       sinput = "<FONT COLOR=#336699 SIZE=1><U>" &  rsx(19) & "</U></FONT>"

    END IF



      celdas = celdas & ofv.GenCelda("","ut ","","3%","","","",sInput,"si")

      sInput = "<b>" & sImagenO & "</b>"
      celdas = celdas & ofv.GenCelda("","ut","","1%","","","",sInput,"si")


      sniv = "---"

                strsql = "Select * From USUARIOS where ID = '" & RSX(13) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)

      sInput = sniv
      celdas = celdas & ofv.GenCelda("","ut" ,"","37%","","","",sInput,"si")

      sInput = "<b>" & sImagenI & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","1%","","","",sInput,"si")

      sniv = "---"

                strsql = "Select * From USUARIOS where ID = '" & RSX(14) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)

      sInput = sniv
      celdas = celdas & ofv.GenCelda("","ut" ,"","37%","","","",sInput,"si")



    filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")




    formu = formu & filas

      celdas = ofv.GenCelda("","AAA COLSPAN=7 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;FONT-SIZE:5pt' ","","1%","","","","&nbsp;","si")

    formu = formu & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF

    posi = posi + 1
    rsX.Movenext
  loop



  GENLISTA = FORMU

END FUNCTION

%>