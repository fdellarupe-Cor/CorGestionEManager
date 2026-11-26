<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%



ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMoverx = request.querystring("CantRegMoverx")
  SWKFCOD     = REQUEST.QUERYSTRING("WKFCOD")
  WKFCODh     = REQUEST.QUERYSTRING("WKFCODh")
  he     = REQUEST.QUERYSTRING("he")

  SOPCION     = REQUEST.QUERYSTRING("OPCION")
'response.write request.querystring     &  session("FBUSR")
  IF SOPCION = "" THEN SOPCION = "NO"


  if SWKFCOD <> "" then
     session("SWKFCOD")     = SWKFCOD
     session("FBUSR")       = SUSUARIO

'''  response.write SUSUARIO & " SUSUARIO linea 25 " & chr(13) & "<br>"

     session("CantRegMoverx")     = CantRegMoverx


  else
     SWKFCOD     = session("SWKFCOD")
 ''  response.write SESSION("FBUSR") & " Session linea 32 " & chr(13) & "<br>"
   IF SESSION("FBUSR") <> " " THEN
   	  IF SESSION("FBUSR") <> "NO" THEN
         SUSUARIOX   = session("FBUSR")
  ''     response.write SUSUARIO & " SUSUARIO linea 34 " & chr(13) & "<br>"
         SUSUARIO    = mid(SUSUARIOX,5,len(SUSUARIOX)-4)
  ''     response.write SUSUARIO & " SUSUARIO linea 34 " & chr(13) & "<br>"
      end if
   end if

     CantRegMoverx     = session("CantRegMoverx")

  end if

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)



  set cn = ofv.conectar(ofv.strconn2)

  IF SOPCION = "REASIGNA" THEN

        USRANT = REQUEST.FORM("RSPANT") & ""
        USRNEW = REQUEST.FORM("RSPNEW") & ""
        CHGOPC = REQUEST.FORM("RSPOPC") & ""

 ''     response.write RSPANT & " RSPANT linea 58 " & chr(13) & "<br>"
 ''     response.write RSPNEW & " RSPNEW linea 59 " & chr(13) & "<br>"
 ''     response.write SUSUARIO & " SUSUARIO linea 60 " & chr(13) & "<br>"
 ''     response.write SUSUARIOX & " SUSUARIOX linea 60 " & chr(13) & "<br>"
 ''     response.write SWKFCOD  & " SWKFCOD linea 61 " & chr(13) & "<br>"

        Strsql = ""                                               ''
        strsql = "UPDATE objprty SET AUTOR = '" & SUSUARIO & "' " ''
        Strsql = Strsql & " WHERE codobjeto = '" & SWKFCOD & "' "   ''
''      response.write Strsql & " Strsql " & chr(13) & "<br>"     ''
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)          ''


  ''''''RESPONSE.REDIRECT "listapendavencer.asp?CANTREGMOVER=" & CANTREGMOVERX

  END IF

  sHTML = ofv.FormHeader(session("FormN"))

  Fila = ""
  xTrColor = "#F9F9F9"
  WKFINSELECT = ""

 ''RESPONSE.WRITE " QS:" & REQUEST.QUERYSTRING & " RF:" & REQUEST.FORM & " FB:" & SESSION("FBUSR") & "<BR>"


  sImagevMa    = "<font face='arial' size=1 xcolor='#aa0000' style='font-weight:900;background-color:#aa0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MUY&nbsp;ALTO&nbsp;</font>"
  sImageva    = "<font face='arial' size=1 xcolor='#f7aa22' style='font-weight:900;background-color:#f7aa22;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ALTO&nbsp;</font>"
  sImagevm    = "<font face='arial' size=1 xcolor='#f7EEDD' style='font-weight:900;background-color:#f7EEDD;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MODERADO&nbsp;</font>"
  sImagevb    = "<font face='arial' size=1 xcolor='#6699CC' style='font-weight:900;background-color:#6699CC;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ESCASO&nbsp;</font>"
  sImagevNUL    = "<font face='arial' size=1 xcolor='#00AA00' style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NULO&nbsp;</font>"

  sImageSI    = "<font face='arial' size=1  style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;SI&nbsp;</font>"
  sImageNO    = "<font face='arial' size=1  style='font-weight:900;background-color:#AA0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NO&nbsp;</font>"

  sImagevn    = "<font face='arial' size=1 xcolor='#dddddd' style='font-weight:900;background-color:#dddddd;color:#ffffff' >----</font>"

  wkstylof = "AAA style='font-family:arial;font-size:6pt;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' "
  wkstylof2 = "UTI style='border-style:solid;border-width:1;border-color:#aaaaaa;' "

''''b.codvinobj = 'VN' AND



  Strsql = ""

  ''   Strsql = Strsql & " select a.* from ( "
  ''   Strsql = Strsql & " SELECT b.*, cast('12/12/2050' as datetime) - (cast(substring(b.segcktpsc,10,5) as integer)  ) as FECCOMIENZO,  "
  ''   Strsql = Strsql & " cast( ( getdate() - ( cast('12/12/2050' as datetime) - (cast(substring(b.segcktpsc,10,5) as integer)  ) )  ) as integer) as DIFER "
  ''   Strsql = Strsql & " FROM CKTVIEW b "
  ''   Strsql = Strsql & " ) A "
  ''   Strsql = Strsql & " WHERE a.segcktpsc = '" & SWKFCOD & "' "
  ''

'' Actualiza responsable
  Strsql = Strsql & " select autor, codobjeto, tipobjeto, fecvto, estado "
  Strsql = Strsql & " from objprty "
  if he="S" then
  Strsql = Strsql & " where "
  else
  Strsql = Strsql & " where estado = '3' and "
  end if
  Strsql = Strsql & " codobjeto = '" & SWKFCOD & "' "

'response.write strsql & "<br>"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

           Celda = ""
           Celda = Celda & ofv.gencelda("","VTG colspan=2 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY DOCUMENTOS A VENCER PARA<BR>LA BUSQUEDA SOLICITADA","si")
           Fila  = FILA & ofv.genrow("","","","","50","","",Celda,"si")


  ELSE
  posI = 0



  wkstylo1 = " style='border-style:solid;border-width:1;border-color:#aaaaaa;COLOR:#000000;' "
  wkstyloT = " style='border-style:solid;border-width:1;border-color:#aaaaaa;COLOR:#FFFFFF;BACKGROUND-COLOR:#888888;' "


        WKIPROCc = 0
        MAINPROC = 0
        MAINACTI = ""

           Celda = ""
           Celda = Celda & ofv.gencelda("","VTI colspan=2 style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"","","","","","LISTA DE DOCUMENTOS A VENCER","si")
           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")


     fi = CLng(CDate("01-01-1900"))

   ''sInput2 = rs("feccomienzo")
     sfecvto = rs(3)
   ''response.write sfecvto & "<br>"
  '' response.write fecvto  & "<br>"

   ''sinput3 = rs("difer")

'''#F7DD88

        WKIACTIV = ""

        WKIPROCD = ""
        WKIPROCN = 0


        Celda = ""

  '' Codigo Objeto
        sInput = RS(1)

        Strsql = ""

     Strsql = ""
     Strsql = Strsql & " select * "
     Strsql = Strsql & " from csrelpathdescriex a   "
     Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto)   "

     Strsql = Strsql & " where  a.cobj = 0" & sInput & "   "

     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    if not rsx1.eof then
     sinput=rsx1("temadescri") & " / " & rsx1("stemadescri")
     SDOC=RSX1("TOBJDESCRI") & " : " & RSX1("EXPR7")
    else
    sinput="---"
    SDOC=""
    end if

   call ofv.cerrarconsulta(rsx1)
   ''''     sInput = WKIPROCD & " - (" & rsX1(0) & ") - " & rs(4)

      Celda = ""
      Celda = Celda & ofv.gencelda("","utI " & wkstyloT,"","30%","","","","TEMA / SUBTEMA:","si")
      Celda = Celda & ofv.gencelda("","UTi XBGCOLOR='#F7EEDD' STYLE='COLOR:#000000;'  " & wkstylo1,"","","","","",sInput,"si")
      Fila  = FILA & ofv.genrow("","aaa Xbgcolor='#ddddff'  ","","","25","","",Celda,"si")

      Celda =  ofv.gencelda("","utI " & wkstyloT,"","30%","","","","DOCUMENTO:","si")
	    Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",SDOC,"si")
      Fila  = FILA & ofv.genrow("","aaa Xbgcolor='#ddddff'  ","","","25","","",Celda,"si")

''''response.write " --> " & SDOC & " sdoc - linea 316 "


''  Autor
    sInput = RS(0)

 '' response.write sInput & " linea 196 " & chr(13) & "<br>"
 ''
 '' xxxx = getInfoUser(RS(0))
 ''
 '' response.write getInfoUser(RS(0)) & " linea 202 " & chr(13) & "<br>"
 ''
 '' response.write session("SGU_nombre") & " linea 204 " & chr(13) & "<br>"
 ''
 ''  if session("SGU_nombre") <> "" then
 ''     sInput =  "<A HREF='WKFREASIGNADAVUSR.ASP' CLASS='NT' TITLE='REASIGNAR PASO A OTRO USUARIO.' ><U>" & session("SGU_nombre") & "</U></A>"
 ''  else
        sInput =  "<A HREF='WKFREASIGNADAVUSR.ASP' CLASS='NT' TITLE='REASIGNAR PASO A OTRO USUARIO.' ><U>" & RS(0) & "</U></A>"
 ''  END IF


      Celda = ""
      Celda = Celda & ofv.gencelda("","uti " & wkstyloT,"","","","","","USUARIO RESPONSABLE:","si")
      Celda = Celda & ofv.gencelda("","UTi  " & wkstylo1 & "  ","","","","","",sInput,"si")
      Fila  = FILA & ofv.genrow("","aaa   ","","","25","","",Celda,"si")



''    Celda = ""
''    Celda = Celda & ofv.gencelda("","utI " & wkstyloT,"","","","","","PASO:","si")
''
''    sInput = RS(7) & "&nbsp;/&nbsp;" & RS(16)
''    Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")
''    Fila  = FILA & ofv.genrow("","aaa   ","","","25","","",Celda,"si")


      Celda = ""
      Celda = Celda & ofv.gencelda("","utI " & wkstyloT,"","","","","","INICIO:","si")


      dfecvto = CDate(CLng(rs(3)) + fi)
      sInput2 = fechadorCor(dfecvto,"")

      If dfecvto = CLng(CDate("01/01/1900")) Then sInput2 = 0

      sinput = sInput2 & "&nbsp;"
      Celda = Celda & ofv.gencelda("", "UTI " & wkstylo1,"","","","","",sInput,"si")
      Fila  = FILA & ofv.genrow("","aaa   ","","","25","","",Celda,"si")



      Celda = ""
      Celda = Celda & ofv.gencelda("","utI " & wkstyloT,"","","","","","NUEVO RESPONSABLE:","si")


      sInput = "SIN DETERMINAR"
      SUSUARIO = "NO"

 '''''response.write Sopcion & " sopcion linea 247 " & chr(13) & "<br>"

   IF SOPCION = "SETUSR" THEN


 ''''''''response.write Sopcion & " sopcion linea 249 " & chr(13) & "<br>"
 ''''''''response.write SESSION("FBUSR") & " FBUSR linea 250 " & chr(13) & "<br>"

         IF SESSION("FBUSR") <> "NO" THEN
            SUSUARIOX = SESSION("FBUSR")
           ' response.write  SESSION("FBUSR") &" 222222 "& mid(SUSUARIOX &"",1,2) & "22222222222222222"

            SUSUARIO  = mid(SUSUARIOX &"",5,len(SUSUARIOX &"")-4)
         END IF

   END IF

''response.write SUSUARIO  & " SUSUARIO linea 259 " & chr(13) & "<br>"
''response.write SUSUARIOX & " SUSUARIOX linea 260 " & chr(13) & "<br>"

      IF SUSUARIO <> "NO" THEN

         xxxx = getInfoUser(SUSUARIO)
'''      response.write SUSUARIO  & " SUSUARIO linea 266 " & chr(13) & "<br>"
         if session("SGU_nombre") <> "" then
            sInput =   "<U><FONT COLOR='#993333' >" & session("SGU_nombre") & "</FONT></U>"
         END IF

      END IF

      Celda = Celda & ofv.gencelda("","UTi  " & wkstylo1 & "  ","","","","","",sInput,"si")

      Fila  = FILA & ofv.genrow("","aaa   ","","","25","","",Celda,"si")

  END IF

  sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa ","","80% ","","0","","0","-1","","",Fila,"si") & "</CENTER>"

''''''response.write SESSION("SETUSR") & " session linea 280 " & chr(13) & "<br>"

      IF SOPCION = "SETUSR"  THEN

''''''   response.write SESSION("FBUSR") & " session linea 283 " & chr(13) & "<br>"

         IF SESSION("FBUSR") <> "NO" THEN
            SUSUARIOX = SESSION("FBUSR")
            'response.write mid(SUSUARIOX,5,len(SUSUARIOX))
            SUSUARIO  = mid(SUSUARIOX,5,len(SUSUARIOX)-4)

'''' response.write susuario  & " reasigna susuario  linea 288 " & chr(13) & "<br>"
'''' response.write susuariox & " reasigna susuariox linea 289 " & chr(13) & "<br>"

            sAccion = "WKFREASIGNADAV.ASP?OPCION=REASIGNA"
            sHTML = sHTML & "<CENTER>"
            sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")
            if he <> "S" then
            sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & RS(0) & "' NAME='RSPANT' >"
            end if
            sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SUSUARIO & "' NAME='RSPNEW' >"

'''''''   response.write RS(0) & " rspant linea 298   " & chr(13) & "<br>"
'''''''   response.write SUSUARIO & " susuario linea 299 " & chr(13) & "<br>"

            'sHTML = sHTML & "<SELECT NAME='RSPOPC' >"
            'sHTML = sHTML & "<OPTION VALUE='1' SELECTED >SOLO PARA ESTE PASO</OPTION>"
            'sHTML = sHTML & "<OPTION VALUE='2'          >CAMBIAR TAMBIEN EN AUTOEVALUACION</OPTION>"
            'sHTML = sHTML & "</SELECT><BR><BR>"


            Celda = ofv.gencelda("","VT  " & wkstylo1 & "   ","CENTER","","","","","<FONT COLOR='#993333' >CONFIRMA LA REASIGNACION DEL RESPONSABLE PARA ESTE PASO???</FONT>","si")

            Fila  = ofv.genrow("","aaa   ","","","35","","",Celda,"si")


            sHTML = sHTML & ofv.GenTabla("","aaa ","","80% ","","0","","0","-1","","",Fila,"si") & "</CENTER>"

           'Botones Aceptar y Cancelar
         sAccion = "listapendavencer.asp?CANTREGMOVER=" & CANTREGMOVERX
         sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")


         END IF

      END IF





  'Botonera
  Agregar = ""
  Buscar  = ""

  sAccion = "listapendavencer.asp?CANTREGMOVER=" & CANTREGMOVERX
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
'  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)

''call ofv.cerrarconn(cnx)
  if he="S"  then
   IF  SUSUARIO <> "NO" THEN
   strsql = "UPDATE OBJSDK SET segown = '" & SUSUARIO & "' " ''
   Strsql = Strsql & " WHERE codobjeto = '" & WKFCODh & "' "   ''
       ' response.write Strsql & " Strsql " & chr(13) & "<br>"     ''
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)
      '  call ofv.cerrarconn(cn)
   END IF
  response.redirect "lisobjhabil.asp"
  end if
  call ofv.cerrarconn(cn)
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

FUNCTION GENCKTMAIL(WETIPO,WECOD,WEMESSAGE,WEORIGEN,WELVL)


    XRESPO = WECOD & ""

    if XRESPO <> "" AND XRESPO <> "NO" then

       set cn2 = ofv.conectar(ofv.strconn0)

       XPARA = GETRESPOCKT(WECOD,WETIPO)
       STRSQL = "SELECT EMAIL,DESCRIPCION FROM USUARIOS WHERE "

     XGUMENSAGEw = WEMESSAGE

     XMESAURL  = ""

     IF  SESSION("URLTYP")  = "V" THEN

         XMESAURL  =  left(request.servervariables("SCRIPT_NAME"),instr(RIGHT(request.servervariables("SCRIPT_NAME"),LEN(request.servervariables("SCRIPT_NAME"))-1),"/"))

     END IF

     IF SESSION("URLPORT")  = "N" THEN
        XMESAURL   = TipoSite("")  & request.servervariables("SERVER_NAME")  & XMESAURL
      ELSE
        XMESAURL   = TipoSite("")  & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
     END IF

    ' XMESAURL   = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL

     XGUMENSAGEw = replace(XGUMENSAGEw,"<APLURL>","<U><A HREF='" & XMESAURL  & "' TARGET='_blank' >")
     XGUMENSAGEw = replace(XGUMENSAGEw,"</APLURL>","</A></U>")

     WEMESSAGE = XGUMENSAGEw


       SELECT CASE WETIPO

        CASE "USR"
           'STRSQL = STRSQL & " ID = '" & XRESPO & "' "
           xxxx = getInfoUserw(XRESPO)
           if session("SGU_mail")<> "" then    SINPUTX= session("SGU_mail")


            stowho = SINPUTX & ""
	   	         if len(stowho) > 0 then
	   	            stowho = trim(stowho)

	   	            subject = "SOLICITUD DE AUTORIZACION"
	   	            IF WETIPO = "USR" THEN
	   	               XPARA2 = XPARA
	   	            ELSE
	   	               XPARA2 = XPARA & " / " & session("SGU_nombre")
	   	            END IF
	   	          ''  message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
                            Xmessage = GenMDLMAIL (XPARA2,WEORIGEN,WEMESSAGE)
	   	            CALL SendMail(Xmessage, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
                          end if


        CASE "GRP"

           'STRSQL = STRSQL & " GRUPO = '" & XRESPO & "' and ( NOT nivel >  " & WELVL & " ) "

           call  getUsrIngroup(XRESPO)
	                  A_SGU_UIGuserid =split(session("SGU_UIGuserid"),chr(9))
	                  A_SGU_UIGdescri =split(session("SGU_UIGdescri") ,chr(9))

	                '   set RSXPA = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	                  ' DO UNTIL RSXPA.EOF
	                   for a_i=0 to ubound(SGU_UIGdescri)-1
	                      xxxx = getInfoUserw(A_SGU_UIGuserid(a_i))
	   		      stowho = session("SGU_mail") & ""
	   		      if len(stowho) > 0 then
	   			 stowho = trim(stowho)

	   			 subject = "SOLICITUD DE AUTORIZACION"
	   			 IF WETIPO = "USR" THEN
	   			    XPARA2 = XPARA
	   			 ELSE
	   			    XPARA2 = XPARA & " / " & A_SGU_UIGdescri(a_i)
	   			 END IF
	   			''' message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
                                 Xmessage = GenMDLMAIL (XPARA2,WEORIGEN,WEMESSAGE)
	   			 CALL SendMail(Xmessage, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
	   		      end if

	   		   'RSXPA.MOVENEXT
		           next



        CASE "SEC"
           XRESPO = 0 & XRESPO

           STRSQL = STRSQL & " SECTOR = " & XRESPO & " and ( NOT nivel >  " & WELVL & " ) "

       END SELECT

    END IF

END FUNCTION

FUNCTION GETRESPOCKT(RESPCOD,RESPTIP)

  GETRESPOCKT = "SIN ASIGNAR"



    XRESPO = RESPCOD & ""

    if XRESPO <> "" AND XRESPO <> "NO" then

       set cn2 = ofv.conectar(ofv.strconn0)
       set cn3 = ofv.conectar(ofv.strconn2)

       SELECT CASE RESPTIP

        CASE "USR"

           xxxx = getInfoUserw(XRESPO)
           sInputx = session("SGU_nombre")



        CASE "GRP"
           strsql = "Select * From GRUPOS where ID = '" & XRESPO & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInputx = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

        CASE "SEC"
           XRESPO = 0 & XRESPO
           strsql = "Select * From ECO where ECO = " & XRESPO
      '''     response.write strsql
           set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
           if not rsniv.eof then sInputx = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

       END SELECT

       GETRESPOCKT = sInputx

    END IF


END FUNCTION


%>