<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTESTRUCTURA.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTGENERAL.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTPROCEDIMIENTOS.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTCONTROLES.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTESQUEMAS.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTTABLAS.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTARCHIVOS.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTSECTORES.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTWEDT.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTOBJDSC.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTFAQ.inc" -->

<%

  sTipo     = request.querystring("Tipo")
  sCodigo   = request.querystring("Codigo")
  sSOLAPA   = request.querystring("SOLAPA")
  sINDICE   = request.querystring("INDICE")

  ncodigo = scodigo


 'if SESSION("VIEWOPC") <> "IXIND" then

'
 '   RESPONSE.REDIRECT "../coress20/corexpress.asp?opt=CC-" & ncodigo & "-" & sTipo & "&NIV1=" & sSOLAPA & "&NIV2=" & sINDICE
'
 'end if

 'usu=lcase("" & request.servervariables("HTTP_USER_AGENT"))

 '  usu=lcase("" & request.servervariables("HTTP_USER_AGENT"))


 ' if SESSION("VIEWOPC") <> "IXIND" then
 '  if instr(usu,"gsa-crawler") =10 then
'
 '    RESPONSE.REDIRECT "../coress20/corexpress.asp?opt=CC-" & ncodigo & "-" & sTipo & "&NIV1=" & sSOLAPA & "&NIV2=" & sINDICE
'
 '  end if
 'session("usuario")="GOOGLE"

usu=lcase("" & request.servervariables("HTTP_REFERER"))
 xentro=false
 if instr(usu,"population.asp") > 0 then
 'response.redirect "http://server-20/corv34des/coress20/corexpress.asp?opt=CC-4775-PCD"
 xentro=true
RESPONSE.REDIRECT "../coress20/corexpress.asp?opt=CC-" & ncodigo & "-" & sTipo & "&NIV1=" & sSOLAPA & "&NIV2=" & sINDICE
 ELSEIF instr(usu,"search.asp") > 0  then

RESPONSE.REDIRECT "../coress20/corexpress.asp?opt=CC-" & ncodigo & "-" & sTipo & "&NIV1=" & sSOLAPA & "&NIV2=" & sINDICE
 end if
session("usuario")="GOOGLE"

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

response.expires = 0




session("ModMain") = "ED"


xapli = 1

  on error resume next



''''''''''''''''''''''''''''




'''  sHTML = ofv.MENUheader("",funcion)



  'sHTML = "<html><body><meta name='robots' content='noarchive'>"
  sHTML = "<html><head>"



  cborderB = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("APLSHW") & ";"
  cborderl = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("APLSHW") & ";"
  borderB = ""
  borderl = ""





   wkproccod = "emdetaproc.asp?Nombre=" & request.querystring("Nombre")
   wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & scodmaster & "&tipmaster=" & stipmaster
   wkproccod = wkproccod & "&CODNIVEL=" & sCODNIVEL
   wkproccod = wkproccod & "&MODAL=" & SMODAL & "&METODO=" & SMETODO

   Session("ASP")     = "vuelveprocesos.asp"
   Session("ASPSTR")  = wkproccod
   Session("ASPSTRD") = Session("ASPSTR")




      set cn = ofv.conectar(ofv.strconn2)

''''''''''''''''''''''''''''''''''''''''''''

    WKPROCDESC =""

      Strsql = "Select A.*,b.descri as tobjdsc From OBJETOS A"
      Strsql = Strsql & " INNER JOIN tipobjeto B ON ( B.tipobjeto = A.tipobjeto )"
      Strsql = Strsql & " where  a.codobjeto = '" & scodigo & "' "



''  response.write strsql & "<br>" & chr(13)
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

      WKPROCDESC =RS("tobjdsc") & ":&nbsp;" & RS(3)

  end if

   call ofv.cerrarconsulta(rs)

 ' WKPROCDESC="&nbsp;"


   Celda = ""
   Celda = "<title>" & WKPROCDESC & "</title><meta name='robots' content='noarchive'></head><body>"
   'Celda = Celda & ofv.GenCelda("","Vt style='Xbackground-color:#993333;Xcolor:#ffffff;color:#990000;' ","","85%","","","",WKPROCDESC,"si") & "<meta name='robots' content='noarchive'></head><body>"

 '  Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  TablaH =  celda ' ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")




   Fila = ""

          wspcolor1  = "#990000"
          wspcolor2  = "#009900"
          sImageO1   = "&#0253;"
          sImageO2   = "&#0254;"
          xTrColor = "#FFFFFF"


  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then
       bVersion = "SI"
    end if
   Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)




 SELECT CASE SSOLAPA

    CASE 1 'propiedades

    CASE 2 'consideraciones

            ID = nCodigo
            Tabla = GENFMTWEDT(nCodigo,STIPO)

    CASE 3 'sectores

            Tabla = GENFMTSECTORES(nCodigo,STIPO)

    CASE 4 'eventos

    CASE 5 'PROCEDIMIENTOS

            nID = sindice

            sCambios = "NO"

            Tabla = GENFMTPROCEDIMIENTOS(nCodigo,STIPO)

    CASE 6 'descripcion


            Tabla = GENFMTOBJDSC(nCodigo,STIPO)

    CASE 7 'controles

            nID = sindice

            sCambios = "NO"

            Tabla = GENFMTCONTROLES(nCodigo,STIPO)

    CASE 8 'esquemas

            nID = sindice

            sCambios = "NO"

            Tabla = GENFMTESQUEMAS(nCodigo,STIPO)

    CASE 9 'tablas

            'Bloque de Defaults
            ncolcant = 1
            nBorde = false
            nblkfc = "#993333"
            nblkbc = "#ffffff"
            nblkff = "ARIAL"
            nblkfs = "10pt"

            nctitfc = "#000077"
            nctitbc = "#ffffff"
            nctitff = "ARIAL"
            nctitfs = "8pt"

            ncconfc = "#336699"
            ncconbc = "#ffffff"
            ncconff = "ARIAL"
            ncconfs = "8pt"
            redim aColumnas(0)

            nStilo = " Style={font-family:" & nblkff & ";font-size:" & nblkfs & ";color:#993333;background-color:#ffffff;font-weight:bold;} "

            nfontx = "</font>"


            nTabla = sindice

            Strsql = "SELECT * From Objtabdef Where objid = " & nTabla
            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

            if  not rs.eof then
                nTitulo = rs(4)
                if nTitulo = "" then nTitulo = "&nbsp;"
                nPerfil = rs(3)
            end if


            Tabla = GENFMTTABLAS(nCodigo,STIPO)

    CASE 10 'documentos

             SKIND = 1


            Tabla = GENFMTARCHIVOS(nCodigo,STIPO)
           ' Tabla = GENFMTARCHIVOS2(nCodigo,STIPO)

    CASE 11 'contenido
            Marca   = "SI"
            id = sindice

'''''''''''''''''''''''''''


            dim Pdef(13,18)

              strsql = "Select * From SysEstprty"
              set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              t = 0
              do until rs.eof
              for ti = 1 to 18
                  if (rs(ti) & "") = "" then
                     sti = ""
                  else
                     sti = rs(ti)
                  end if

                  Pdef(t,ti - 1) = sti

              next
              rs.movenext
              t = t + 1
              loop

              call ofv.cerrarconsulta(rs)

              strsql = "Select * From Objestr Where objid = " & ID
             ''' response.write strsql & "<br>"
              set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              bsecuen = ""
              bDescri = " "
              bobj = ""
              if not rs.eof then
                 bsecuen = rs(4)
                 bDescri = rs(3)
                 bobj = rs(0)

                 stexto1 = "<A ID='ATIT" & rs(0) & "' NAME='ATIT" & rs(0) & "' "  & chr(13)
                 stexto1 = stexto1 & " class=nt href='coestrCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
                 stexto1 = stexto1 & "&sdocum=" & sdocum & "&sdocc=" & rs(0)
                 stexto1 = stexto1 & "'   >" & sfidkc & "&nbsp;<big>" & rs(3) & "</big></a>" & chr(13)

                 Celdat = "<td class=ut style='cursor:pointer;'  onclick='document.all.ATIT" & rs(0) & ".click();' >" & stexto1 & "</td>"
                 Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si") & chr(13)


              end if
              call ofv.cerrarconsulta(rs)



              strsql = "Select * From ObjEstPrty where detcon = " & ID
              set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              t = 0
              do until rs.eof
              select case rs(4)
                         case "A"
                              t = 0
                         case "B"
                              t = 1
                         case "C"
                              t = 2
                         case "D"
                              t = 3
                         case "E"
                              t = 4
                         case "F"
                              t = 5
                         case "G"
                              t = 6
               end select

              for ti = 5 to 22
              if (rs(ti) & "") = "" then
                 sti = ""
              else
                 sti            = rs(ti)
                 Pdef(t,ti - 5) = sti
              end if



              next
              rs.movenext
              loop

              call ofv.cerrarconsulta(rs)

              PInden = Pdef(0,7)
              if PInden <> "NO" then PInden = "SI"

              PNumer = Pdef(0,12)
              if PNumer <> "NO" then PNumer = "SI"

              PNumCls = Pdef(0,14)
              redim preserve PNumTip(13)
              for i=0 to 13
                  PNumTip(i) = Pdef(i,13)
                  if PNumTip(i) = "" then PNumTip(i) = "1"
              next

'''''''''''''''''''''''''''
            Tabla = GENFMTESTRUCTURA(nCodigo,STIPO)

    CASE 12 'cursograma

    CASE 13 'contratos

             SKIND = 2


             Tabla = GENFMTARCHIVOS(nCodigo,STIPO)

    CASE 14 'faq

             Tabla = GENFMTFAQ(nCodigo,STIPO)

    CASE 15 'general



              Tabla = GENFMTGENERAL(nCodigo,STIPO)



 END SELECT





  call ofv.cerrarconn(cn)

'''''''''''''''''''''''''''

  Celda = TablaH
  sHTML = sHTML & Celda

  Celda = Tabla
  sHTML = sHTML & Celda


  sHTML = sHTML & "</body></html>"

 END IF

  response.write  sHTML
''------------------

FUNCTION GENFMTARCHIVOS2(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""


  strsqlx = "Select Version,estado,tipobjeto From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
      Nestado=rs(1)

  end if
  call ofv.cerrarconsulta(rs)


    sarch = "NO"

    Strsql = "SELECT descri, archivo, objid, directorio "
    Strsql = Strsql & " from ObjHtm Where codobjeto = " & nCodigo '& " AND TIPO ='" & SKIND &"' 	"
    Strsql = Strsql & " order by secuen "

  Formulario = ""
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    do until rs.eof

    sdirec = rs(3) & ""

    if sdirec = "" then sdirec = "Archivos"




       sarch = sdirec & "/" & rs(1)
       clase = "<font color=#993333><u>"
       stexto = sfida & sfta & replace(rs(0)," ","&nbsp;") & sftc

    stexto1 = "<A class=ttt  href='" & sarch&"'"

    stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"

    if (rs(1) & "") = "" then
       stexto1 = "<A class=ttt  href='javascript:void(0);' "
       stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"
    end if

       Celda = ""


      ' estado observacion

      Celdaobs  = ofv.GenCelda("","ut","","1%","","","","&nbsp;","si")




      if Nestado=5  then


               sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(2)  & "&ID2=" & rs(2)

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")
               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")

''               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")
              'Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")


      end if


      Celda = Celda & Celdaobs


       Celda = Celda & ofv.GenCelda("","ut  ","","95%","","","",stexto1,"si")
       Formulario = Celda
       Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")

    rs.movenext
    loop

  FNSHTML = FNSHTML & "<br><br><center>" & ofv.GenTabla("","aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-COLOR:#dddddd;' ","","90%","","0","0","4","0","","",Fila,"si")
 '' if not sdocum = "" then
  if sarch <> "NO" then
     sAccion = "CoArchivos3.asp?Archivo=" & sarch
     FNSHTML = FNSHTML & ofv.GenForm("cons","",sAccion,"_blank","post","",sInput,"si")
  end if

  else

    sdocus = "NO"


  end if

'''SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where codcomp

    fila = ""

    sarch = 0
    'strsql = "SELECT DISTINCT COMPLEMENTOS.codcomp, COMPLEMENTOS.codiden, "
    'strsql = strsql & "COMPLEMENTOS.descri, CSACOMPLE.nombre FROM COMPLEMENTOS, "
    'strsql = strsql & " OBJACCION, CSACOMPLE where COMPLEMENTOS.codcomp = OBJACCION.codcomp "
   ' strsql = strsql & " and COMPLEMENTOS.codcomp = CSACOMPLE.codcomp "
    'strsql = strsql & " and OBJACCION.codobjeto =" & ncodigo
    'strsql = strsql & " ORDER BY COMPLEMENTOS.codcomp "


      Strsql = "Select distinct a.codobjetol, b.codiden, b.nombre , b.nombre From objlnk a, csacomple b  Where a.ltipo = 'KL' and a.codobjeto = 0" & ncodigo
	  strsql = strsql & " and a.codobjetol = b.codcomp "
      strsql = strsql & " ORDER BY a.codobjetol"
      'response.write strsql
    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
  Formulario = ""
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    do until rs.eof


   if sdocum = cstr(rs(0)) then
       sarch = rs(0)
       clase = "<font color=#993333><u>"
       stexto = sfida & sfta & replace(rs(3) & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(2)," ","&nbsp;") & sftc
    else
       clase = ""
       stexto = sfidc & sft & replace(rs(3) & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(2)," ","&nbsp;")
    end if
    stexto1 = "<A class=ttt  href=CoArchCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
    stexto1 = stexto1 & "&sdocum=" & rs(0)
    stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"


       Celda = ofv.GenCelda("","ut  ","","1%","","","",stexto1,"si")
       Formulario = Celda

       rs.movenext

       Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")
    loop

    FNSHTML = FNSHTML & "<br><br><center>" & ofv.GenTabla("","aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-COLOR:#dddddd;' ","","90%","","0","0","4","0","","",Fila,"si")

    if  sarch > 0 then
        sAccion = "EditarImagen2.asp?ID=" & sarch
        FNSHTML   = FNSHTML & ofv.GenForm("cons","",sAccion,"_blank","post","",sInput,"si")
    end if

  else
    scomple = "NO"
  end if
    Call OFV.cerrarconsulta(rs)


  if sdocum <> "" then
     FNSHTML   = FNSHTML & "<script language=javascript >"
     FNSHTML   = FNSHTML & " document.cons.submit(); "
     FNSHTML   = FNSHTML & "</script>"
  end if




  if scomple = "NO" and sdocus = "NO" then
    sestr = "<h5>Sin&nbsp;datos&nbsp;disponibles"
    sestr = sestr & "</h5>"
    Formulario = ofv.GenCelda("","ttt","center","80%","","","",sestr,"si")
    Fila = ofv.GenRow("","","","","","","",Formulario,"si")
    FNSHTML = FNSHTML & "<br><br><center>" & ofv.GenTabla("","aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-COLOR:#dddddd;' ","","90%","","0","0","4","0","","",Fila,"si")
  end if


  GENFMTARCHIVOS2 = FNSHTML


 END FUNCTION



%>

