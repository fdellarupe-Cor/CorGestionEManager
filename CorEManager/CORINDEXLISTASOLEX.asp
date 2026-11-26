<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%
session("usuario")="GOOGLE"
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

response.expires = 0

    sTipo     = request.querystring("Tipo")
    sCodigo   = request.querystring("Codigo")
    sSOLAPA   = request.querystring("SOLAPA")
  sINDICE   = request.querystring("INDICE")


 '  usu=lcase("" & request.servervariables("HTTP_USER_AGENT"))

  ' RESPONSE.WRITE XUS1 & "--" & XUS2

' if SESSION("VIEWOPC") <> "IXIND" then
 ' if instr(usu,"gsa-crawler") = 10 then

 '   RESPONSE.REDIRECT "../coress20/corexpress.asp?opt=CC-" & sCodigo & "-" & sTipo '& "&NIV1=" & sSOLAPA & "&NIV2=" & sINDICE
'
 ' end if
 usu=lcase("" & request.servervariables("HTTP_REFERER"))
 xentro=false
 if instr(usu,"population.asp") > 0 then
 'response.redirect "http://server-20/corv34des/coress20/corexpress.asp?opt=CC-4775-PCD"
 xentro=true
 RESPONSE.REDIRECT "../coress20/corexpress.asp?opt=CC-" & ncodigo & "-" & sTipo & "&NIV1=" & sSOLAPA & "&NIV2=" & sINDICE
 ELSEIF instr(usu,"search.asp") > 0  then

RESPONSE.REDIRECT "../coress20/corexpress.asp?opt=CC-" & ncodigo & "-" & sTipo & "&NIV1=" & sSOLAPA & "&NIV2=" & sINDICE
 end if




session("ModMain") = "ED"


xapli = 1

  on error resume next



''''''''''''''''''''''''''''




''  sHTML = ofv.MENUheader("",funcion)

  'sHTML = "<html><body><meta name='robots' content='noarchive'>"
  sHTML = "<html><head>"'<meta name='robots' content='noarchive'></head><body>"


'''  sHTML = sHTML & ofv.CorFormHeader


  dim stabla(20)
  dim solord(20)
  dim Nombres(20)
  dim ASPs(20)
  dim Solapas(20)
  dim Solapas2(20)

  solord(0) = 0
  solord(1) = 15
  solord(2) = 1
  solord(3) = 6
  solord(4) = 3
  solord(5) = 2
  solord(6) = 4
  solord(7) = 5
  solord(8) = 7
  solord(9) = 8
  solord(10) = 11
  solord(11) = 10
  solord(12) = 9
  solord(13) = 12
  solord(14) = 13
  solord(15) = 14
  solord(16) = 16
  solord(17) = 17
  solord(18) = 18
  solord(19) = 19
  solord(20) = 20


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

    WKPROCDESC = ""

      Strsql = "Select A.*,b.descri as tobjdsc From OBJETOS A"
      Strsql = Strsql & " INNER JOIN tipobjeto B ON ( B.tipobjeto = A.tipobjeto )"
      Strsql = Strsql & " where  a.codobjeto = '" & scodigo & "' "



''  response.write strsql & "<br>" & chr(13)
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

      WKPROCDESC = RS("tobjdsc") & ":&nbsp;" & RS(3)

  end if



   'WKPROCDESC="&nbsp;"


   Celda = ""
   Celda =   "<title>" & WKPROCDESC & "</title><meta name='robots' content='noarchive'></head><body>" 'Celda & ofv.GenCelda("","Vt style='Xbackground-color:#993333;Xcolor:#ffffff;color:#990000;' ","","85%","","","",WKPROCDESC,"si")

   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   TablaH = celda 'ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")




   Fila = ""

          wspcolor1  = "#990000"
          wspcolor2  = "#009900"
          sImageO1   = "&#0253;"
          sImageO2   = "&#0254;"
          xTrColor = "#FFFFFF"


sFuncion = "C"





  perfilsp = "S"


  siconmas = "ICorAcciones.jpg"
  Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sEsquema = rs(1)
     sPerfil = rs(3)
  end if

  call ofv.cerrarconsulta(rs)


  for z=0 to 20
      Solapas(z) = "N"
  next




  Strsql = "Select * From Objperfildoc Where codobjeto =" & sCodigo & " and tipobjeto ='" & stipo & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     perfilsp = "N"
  end if

  if perfilsp = "N" then
     call ofv.cerrarconsulta(rs)
     Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  end if



    for xx = 1 to 20
        solord(xx) = 17
        Nombres(xx) = "N"
    next



    for xx = 2 to 16
        stabla (xx - 1) = rs(xx + 30)
    next



    yy = 0
    for xx = 1 to 15

        if cint(stabla (xx)) < 99 then
           yy = yy + 1
           solord (stabla (xx)) = xx
        end if
    next
    tyy = yy




     xstrsol     = 1






  X=0
  for z = 2 to 16

    if rs(z) Then

      if (sFuncion = "E"  ) or (sFuncion = "C" and z <> 2 and z <> 4 and z <> 5 ) or (sFuncion = "CD" and z <> 2 and z <> 5 )  or (sFuncion = "D"  and z = 4 ) then
          x = x + 1
          Solapas(z-1)  = REPLACE(RS(z + 15)," ","*")
          Solapas2(z-1) = ASPs(z-1)

          if sFuncion <> "E" and z = 13 then
             Solapas(z-1)  = "N"
          end if


          if sFuncion <> "E" and z = 66 then
             Solapas(z-1)  = "N"
             Solapas(16)  = "CIRCUITO COMPLETO"
             Solapas2(16) = ASPs(16)

             for xiord=0 to 19

                 if solord(xiord) = 5 then

                    xmovey = 16

                    for xiord2=xiord + 1 to 19
                        xmoven           = solord(xiord2)
                        solord(xiord2)   = xmovey
                        xmovey           = xmoven
                    next



                    exit for
                 end if
             next


          end if

      end if
    end if

  next


  xfinal = ubound(Solapas)

  call ofv.cerrarconsulta(rs)






  xSolasp = ""
  xSolasp = xSolasp & "BarraSol.asp?Tipo=" & sTipo & "&Codigo="
  xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
  xSolasp = xSolasp & "&Funcion=" & sFuncion
  xSolasp = xSolasp & "&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
  xSolasp = xSolasp & "&SSCLICK="

 r1 = ""






  for x = xstrsol to xfinal
      if solapas(solord(x)) <> "N" then
         sSolapa =  solapas(solord(x))
         sInput  = replace(Ucase(sSolapa),"*","&nbsp;")

         XGUTABSOLAPA = solord(x)
         IF  solord(x) = 16 THEN XGUTABSOLAPA = 5

         CALL GETLISTNV2("L",SCODIGO,STIPO,XGUTABSOLAPA,"0")
         XGUTABKEY = SPLIT(XGUTABLISKEY,CHR(9))
         XGUTABDAT = SPLIT(XGUTABLISDAT,CHR(9))

         XGUTABCANT = UBOUND(XGUTABKEY)

         IF  XGUTABCANT > 0 THEN


                FOR XGUINX=1 TO UBOUND(XGUTABKEY)


                    sInput2  = sInput & "&nbsp;-&nbsp;" & XGUTABDAT(XGUINX)


                    wkproccod = "<a class=nt HREF='CORINDEXLISTADETEX.ASP?CODIGO=" & SCODIGO & "&TIPO=" & STIPO  & "&SOLAPA=" & solord(x) & "&INDICE="   &XGUTABKEY(XGUINX) & "' "
                    wkproccod = wkproccod & " target='_blank' >"
                    wkproccod = wkproccod & sInput2
                    wkproccod = wkproccod & "</a>"



                    Celda     = ""
                    Celda = Celda & ofv.GenCelda("","Ut","","90%","","","",wkproccod,"si")

                    FilaN = filaN & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")


                    if xTrColor = "#FFFFFF" then
                       xTrColor = "#F0F0F0"
                    ELSE
                       xTrColor = "#FFFFFF"
                    END IF


                NEXT

            ELSE


                    sInput2  = sInput


                    wkproccod = "<a class=nt HREF='CORINDEXLISTADETEX.ASP?CODIGO=" & SCODIGO & "&TIPO=" & STIPO  & "&SOLAPA=" & solord(x) & "&INDEX=0" & "' "
                    wkproccod = wkproccod & " target='_blank'  >"
                    wkproccod = wkproccod & sInput2
                    wkproccod = wkproccod & "</a>"



                    Celda     = ""
                    Celda = Celda & ofv.GenCelda("","Ut","","90%","","","",wkproccod,"si")

                    FilaN = filaN & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")


                    if xTrColor = "#FFFFFF" then
                       xTrColor = "#F0F0F0"
                    ELSE
                       xTrColor = "#FFFFFF"
                    END IF


            END IF




      end if
  next



       Tabla = "<BR>" & Tabla & ofv.GenTabla("","","","99%","","0","","1","2","","",Filan,"si")












  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'''''''''''''''''''''''''''

  Celda = TablaH
  sHTML = sHTML & Celda

  Celda = Tabla
  sHTML = sHTML & Celda


  sHTML = sHTML & "</body></html>"

 END IF

  response.write  sHTML



%>

