<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
    response.expires=0

    session("noved") = "SI"
    stilo1 = " style={background-color:" & session("CMFMENU") & ";border-style:outset;border-width:2;}"
    stilo2 = " style={background-color:" & session("CMFMNBOT") & ";}"


    set cn = ofv.conectar(ofv.strconn2)
    set cn0 = ofv.conectar(ofv.strconn0)

dim atit(10)
titc = 0
Strsql = "SELECT * FROM novedtit "
set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if not rsx.eof then

   for tt = 2 to 4
       atit(tt) = rsx(tt)
       if rsx(tt) <> "NO" then titc = titc + 1
   next


end if
call ofv.cerrarconsulta(rsx)

menuSel = request.querystring("menu")
menu    = menuSel
if menu = "" then menu = 3

menu    = cint(menu)

  if session("FormN") = "" then session("FormN") = "NOVEDADES"

  sHTML = ofv.FormHeader(session("FormN"))

 '' sHTML = ofv.menuheader("","novedades ")

    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos,nom) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

      s =  "<script language=JavaScript>" & Chr(13)
      s = s & "function AbrirVentana(Datosx) {" & Chr(13)
      s = s & " myWin= open(Datosx"
      s = s & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
      s = s & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
      s = s & " ); " & Chr(13) & "window.close(); " & Chr(13)
      s = s & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs & s









nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

   dFechai  = clng(date()) - nFecini
   XTrColor = "#FFFFFF"

   sglmar   = "<font class=btiW >*</font>"
   sglman   = "<font class=btiW >1</font>"
   sglprf   = "<font class=btic >&#0128;</font>"
   sgllnk   = "<font class=btic >&#0126;</font>"
   sgladj   = "<font class=btic >&#0204;</font>"


   dim ANTabla(1000,8)
   TNTabla = 0
   INTabla = 0
   xnohay1 = "S"
   xnohay2 = "S"
   SaleSel = "N"

Do until SaleSel = "S"

   select case menu

       case 1


       case 2
				'RESPONSE.WRITE(CHR(13) & "CASE 2" & CHR(13))
                         session("ModMain") = "CL"

                         Strsql = "SELECT * FROM csnovedades "
                         Strsql = Strsql & " Where fecexp > " & dFechai
                         Strsql = Strsql & " and fecinicio < " & dFechai + 1
                         Strsql = Strsql & " And ("
''INICIO CODIGO ANTERIOR*****
'Strsql = Strsql & " (tipo = 2 And usugrp = '" & session("usuario") & "') or "
''FIN CODIGO ANTERIOR
''*****INICIO COPIADO DE GP
                         Strsql = Strsql & " (tipo = 2 And usugrp = '" & session("usuario") & "') or contacto ='" & session("Usuario") & "' or"
''*****FIN COPIADO DE GP

                         Strsql = Strsql & " (tipo = 1 And usugrp = 'TODOS' and sector = 999 ) "
''*****INICIO CODIGO ANTERIOR*****
                         'Strsql = Strsql & ") "
''*****FIN CODIGO ANTERIOR*****

''*****INICIO COPIADO DE GP*****
Strsql = Strsql & ") " 'and contacto='" & session("Usuario") & "'"
''*****FIN COPIADO DE GP*****
                         Strsql = Strsql & " order by fecinicio desc, tipo desc, importancia, objid desc"
                        ' response.write strsql
                         set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         if   rsx.eof then
                              xnohay1 = "S"
                         else
                              xnohay1 = "N"
                           do until rsx.eof
                              INTabla = INTabla + 1

                              xfromusr = rsx(11)
                              strsql = "Select descripcion From Usuarios  where ID = '" & rsx(11) & "' "
                              set rsx1 = ofv.crearconsultaEx(StrSql,cn0,1,volver)
                              if  not rsx1.eof then
                                  xfromusr = rsx1(0)
                              end if
                              call ofv.cerrarconsulta(rsx1)

                              ANTabla(INTabla,0) = rsx(1)
                              ANTabla(INTabla,1) = clng(rsx(2))
                              ANTabla(INTabla,2) = clng(rsx(6))
                              ANTabla(INTabla,3) = clng(rsx(3))
                              ANTabla(INTabla,4) = rsx(12)
                              ANTabla(INTabla,5) = "N"
                              ANTabla(INTabla,6) = "N"
                              ANTabla(INTabla,7) = xfromusr
                              ANTabla(INTabla,8) = rsx(5)

                              rsx.movenext
                            loop
                        end if
                        call ofv.cerrarconsulta(rsx)

       case 3
				'RESPONSE.WRITE(CHR(13) & "CASE 3" & CHR(13))
                        session("ModMain") = "CL"

                        Strsql = "SELECT * FROM csnovedades "
                        Strsql = Strsql & " Where fecexp > " & dFechai
                        Strsql = Strsql & " and fecinicio < " & dFechai + 1
                        Strsql = Strsql & " And tipo = 3"
                        Strsql = Strsql & " And ((usugrp = '" & session("usuario") & "') "
                        Strsql = Strsql & " or ( usugrp = 'TODOS' and sector = 999))  order by fecinicio desc, importancia, objid desc"

						'response.write (chr(13) & strsql & chr(13))

                        set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                        if   rsx.eof then
                             xnohay1 = "S"

                        else
                             xnohay1 = "N"
                             do until rsx.eof

                                INTabla = INTabla + 1

                                Strsql = "SELECT archivo FROM tipobjeto "
                                Strsql = Strsql & " Where tipobjeto = '" & rsx(9) & "'"
                                set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                if  not rs.eof then
                                    'xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
									xarc = ""
                                else
                                    xarc = "NO"
                                end if
                                call ofv.cerrarconsulta(rs)


                                sActcon = "Consultar.asp?Codigo=" & rsx(8) & "&Tipo=" & rsx(9)
                                sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
                                sHref = "&nbsp;"
                              ' response.write sActcon
                                sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)


                                wmsgx = rsx(5) & ""


                                if len(wmsgx) < 5 then
                                   wmsgx = esqnomP(rsx(9),rsx(8),ofv) & " VIGENTE A PARTIR DEL " & fechaArg(cdate(clng(rsx(3)) + clng(cdate("01-01-1900")))) & ". "
                                else
'                                  response.write "aa" & len(wmsgx) & rsx(5)
                                   wmsgx = Ucase(wmsgx)
                                end if

                                wNomL = "&nbsp;" & wmsgx & "&nbsp;"



                                sImagen = wNomL


                                if xarc <> "NO" then
                                   sImagen = xarc & "&nbsp;" & sImagen
                                end if

                                sLink = OFV.GenLink("", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")


                                xfromusr = rsx(11)
                                strsql   = "Select descripcion From Usuarios  where ID = '" & rsx(11) & "' "
                                set rsx1 = ofv.crearconsultaEx(StrSql,cn0,1,volver)
                                if  not rsx1.eof then
                                    xfromusr = rsx1(0)
                                end if
                                call ofv.cerrarconsulta(rsx1)

                              ANTabla(INTabla,0) = rsx(1)
                              ANTabla(INTabla,1) = 3
                              ANTabla(INTabla,2) = clng(rsx(6))
                              ANTabla(INTabla,3) = clng(rsx(3))
                              ANTabla(INTabla,4) = rsx(12)
                              ANTabla(INTabla,5) = "S"
                              ANTabla(INTabla,6) = "N"
                              ANTabla(INTabla,7) = xfromusr
                              ANTabla(INTabla,8) = sLink

                            rsx.movenext
                            loop
                        end if
                        call ofv.cerrarconsulta(rsx)

        case 4
			'RESPONSE.WRITE(CHR(13) & "CASE 4" & CHR(13))

                        session("ModMain") = "WK"
                        Session("Segce")   = "S"
                        Session("Segco")   = "S"

                        Strsql = "SELECT * FROM csnovedades "
                        Strsql = Strsql & " Where fecexp > " & dFechai
                        Strsql = Strsql & " and fecinicio < " & dFechai + 1
                        Strsql = Strsql & " And tipo = 4"
                        Strsql = Strsql & " And usugrp = '" & session("usuario") & "' order by fecinicio desc, importancia, objid desc"

						'response.write(chr(13) & strsql & chr(13))

                        set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                        if  rsx.eof then
                            xnohay1 = "S"
                        else
                            xnohay1 = "N"
                            do  until rsx.eof


                            INTabla = INTabla + 1

                            Strsql = "SELECT archivo FROM tipobjeto "
                            Strsql = Strsql & " Where tipobjeto = '" & rsx(9) & "'"
                            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                            if  not rs.eof then
                                'xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
								xarc = ""
                            else
                                xarc = "NO"
                            end if
                            call ofv.cerrarconsulta(rs)


                            Session("ASP")      = "novedmenu.asp?menu=4"
                            session("sEstado")  = "PARA APROBAR"
                            session("sEstadon") = "2"
                            Session("Form2")    = "apr"

                            aspeditar = "Editar.asp?Codigo=" & rsx(8) & "&Tipo=" & rsx(9) & "&Perfil=xxxx&Funcion=E"
                            aspeditar = cstr(aspeditar)
                            sHref     = "&nbsp;"
                          ' response.write sActcon
                            sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                            wNomL = "&nbsp;" & Ucase(rsx(5)) & "&nbsp;"

                            sImagen = wNomL

                            if  xarc <> "NO" then
                                sImagen = xarc & "&nbsp;" & sImagen
                            end if

                            sLink = OFV.GenLink("", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")


                            xfromusr = rsx(11)
                            strsql   = "Select descripcion From Usuarios  where ID = '" & rsx(11) & "' "
                            set rsx1 = ofv.crearconsultaEx(StrSql,cn0,1,volver)
                            if  not rsx1.eof then
                                xfromusr = rsx1(0)
                            end if
                            call ofv.cerrarconsulta(rsx1)


                            ANTabla(INTabla,0) = rsx(1)
                            ANTabla(INTabla,1) = 4
                            ANTabla(INTabla,2) = clng(rsx(6))
                            ANTabla(INTabla,3) = clng(rsx(3))
                            ANTabla(INTabla,4) = rsx(12)
                            ANTabla(INTabla,5) = "S"
                            ANTabla(INTabla,6) = "N"
                            ANTabla(INTabla,7) = xfromusr
                            ANTabla(INTabla,8) = sLink



                         rsx.movenext
                         loop
                        end if
                        call ofv.cerrarconsulta(rsx)

'''''''''''''


set cn1 = ofv.conectar(ofv.strconn5)
set cn4 = ofv.conectar(ofv.strconn6)

                         'xtipos = " OBJTIPO in ('RSK'"
                         'separ = ","
						 xtipos = " OBJTIPO in ( "
                         Strsql = "SELECT tipobjeto FROM tipobjeto where aplic in ('EM','GLB')"
                         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         if  not rs.eof then
                             do while not rs.eof
                                xtipos = xtipos & separ & "'" & rs(0) & "'"
                                separ = ","
                                rs.movenext
                             loop
                         else
                             xtipos = xtipos & "' '"
                         end if
                         xtipos = xtipos & " )"

                         SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & Session("Usuario") & "') OR ("

                         xxxx     = getInfoUser(Session("Usuario"))
                         xusrlog  = session("SGU_nombre")



                         if  xusrlog <> "NO" THEN


                             SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "') ) "
                             SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

'''''''''''''''''''''''''''''''''''''''''''
                             if  Session("DELLOG") = "S" then

                                 XMFTGRP = split(session("SegGrpLst"),chr(9))
                                 xcount  = ubound(XMFTGRP) + 1
                                 XMFSEP  = ""
                                 SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

                                 for XMFI = 0 to ubound(XMFTGRP)
                                     SEGWHERE = SEGWHERE  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                                     XMFSEP  = ","
                                 next

                                 SEGWHERE = SEGWHERE & ") ) )"

                            else
                                 SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
                            end if


''''''''''''''''''''''''''''''''''''''''''

                          '''   IF  RS(21) > 0 THEN


                             if  cstr(session("SGU_idcostcenter") & "") <> "0" then
                                 SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
                                 SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & session("SGU_idcostcenter") & "' OR SUBCODIGO = '" & session("SGU_idcostcenter") & "' ) "
                                 SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
                                 SEGWHERE = SEGWHERE & " ) ) "
                             END IF


                         END IF

                         SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "

                         Strsql = "SELECT * FROM CKTVIEW " & SEGWHERE

						'RESPONSE.WRITE chr(13) & STRSQL & chr(13)

                         set rsx = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
                         if  rsx.eof then
                             xnohay2 = "S"
                         else
                             xnohay2 = "N"
                             do  until rsx.eof


								'RESPONSE.WRITE chr(13) & rsx(0) & "  " & rsx(1) & "  " & rsx(2) & "  " & rsx(3) & chr(13)


							xvale = "N"

							if rsx(3) = "RSK" then

								Strsql  = "SELECT * FROM ORRSKCIRCUITO "
								Strsql  = Strsql & " Where OBJID = 0" & rsx(4)
								set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								if  not rsxp.eof then
									xvale = "S"
									xMAINAPL = rsxp("APLIC")
								end if
								call ofv.cerrarconsulta(rsxp)


							else



								Strsql  = "SELECT * FROM objetos "
								Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
								set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
								if  not rsxp.eof then
									xvale = "S"
									xMAINAPL = rsxp("APLIC")
								end if
								call ofv.cerrarconsulta(rsxp)

							end if




								 IF  xvale = "S" THEN

                                 xnohay2 = "N"
                                  INTabla = INTabla + 1
                                '' xarc = "<font face=corflaticon color=#6699cc size=1 style='font-size:9pt;line-height:10pt;border-style:outset;border-width:1;border-color:#6699cc;' >&#0210;</font>"

                                 Session("ASP")      = "novedmenu.asp?menu=4"
                                 session("sEstado")  = "PARA APROBAR"
                                 session("sEstadon") = "2"
                                 Session("Form2")    = "apr"

                                 aspeditar = "Editar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=E&SEGCKTPSC=" & RSX(8)
                                 aspeditar = cstr(aspeditar)
                                 sHref     = "&nbsp;"
                        '        response.write sActcon
                                 sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                                 wNomL     = "&nbsp;CIRCUITO:&nbsp;" & RSX(7) & "&nbsp;&nbsp;&nbsp;PASO:&nbsp;"  & RSX(16) & "&nbsp;"

                                 sImagen   = wNomL

                                 if xarc <> "NO" then
                                    sImagen = xarc & "&nbsp;&nbsp;" & EsqnomP(rsx(3),rsx(4),ofv) & "&nbsp;&nbsp;" & sImagen
                                 end if

                                 sLink     = OFV.GenLink("", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")

                                 xfromusr = rsx(6)
                                 xxxx     = getInfoUser(xfromusr)
                                 xusrlog  = session("SGU_nombre")

                                 ANTabla(INTabla,0) = "CKT" & INTabla
                                 ANTabla(INTabla,1) = 5
                                 ANTabla(INTabla,2) = 0
                                 ANTabla(INTabla,3) = dFechai
                                 ANTabla(INTabla,4) = rsx(10)
                                 ANTabla(INTabla,5) = "S"
                                 ANTabla(INTabla,6) = "N"
                                 ANTabla(INTabla,7) = xusrlog
                                 ANTabla(INTabla,8) = sLink
                                 end if
                             rsx.movenext
                             loop
                        end if
                        call ofv.cerrarconsulta(rsx)



       case else


     end select

     if  menuSel <> "" then
         saleSel = "S"
     else
         if  INTabla > 0 then
             saleSel = "S"
         else
             menuSel = 2
             menu    = 2
         end if
     end if



  loop

     call ofv.cerrarconn(cn)
     call ofv.cerrarconn(cn0)

     TNTabla = INTabla

     sHTML   = sHTML &  "<table class='atbh' width=100% cellspacing=0 cellpadding=2  >"

     sHTML   = sHTML &  "<tr>"

   '''  xmDato  = GenTitulo(atit(menu),0)

     xmDato  = "<u>" & atit(menu) & "</u>"

     sHTML   = sHTML &  "<td class='ttr' width=2%  ><big><font class=btic >4</font></big></TD>"
     sHTML   = sHTML &  "<td class='ttr' width=30%  >"  & xmDato & "</TD>"

     sHTML   = sHTML &  "<td class='ttr' width=70%  align=right >"

             sHTML = sHTML &  "<table  cellspacing=2 cellpadding=2  >"

             for ii=2 to 10
                 if  atit(ii) <> "NO" then

                    if  ii = menu then
                        stilo  = stilo2
                    else
                        stilo  = ""
                     ''   xmDato = GenTitulo(atit(ii),1)
                        xmDato  = atit(ii)
                        sHTML  = sHTML &  "<td class=ttr align=center style={border-style:solid;border-width:2;"
                        sHTML  = sHTML &  "border-color:#4d4d4d} >"
                        sHTML  = sHTML &  "<a style='color:#4d4d4d;' href=NOVEDADM.asp?menu=" & ii & ">" & xmDato & "</a></TD>"
                    end if

                 end if
            next

            sHTML = sHTML &  "</table>"

  sHTML = sHTML &  "</TD>"
  sHTML = sHTML &  "</TR>"
  sHTML = sHTML &  "</table>"


  sHTML = sHTML &  "<center>"



  if  xnohay1 = "S" and xnohay2 = "S" then

      sHTML = sHTML &  "<table  margin=3 width=95% cellspacing=0 cellpadding=0  style='position:relative;top:+5;border-style:solid;border-width:1;border-color:#eeeeee;'  >"
      sHTML = sHTML &  "<tr  height=200 >"
      sHTML = sHTML &  "<td class=ttt width=100% align=center valign=middle ><big>SIN&nbsp;NOVEDADES&nbsp;DE&nbsp;" & atit(menu) & "&nbsp;</TD>"
      sHTML = sHTML &  "</TR>"

  else
      sHTML = sHTML &  "<table bgcolor=#eeeeee margin=3 width=95% cellspacing=0 cellpadding=0  style='position:relative;top:+5;border-style:solid;border-width:1;border-color:#eeeeee;'  >"

      for INTabla=1 to TNTabla

          xtmcod   = ANTabla(INTabla,0)
          xtipom   = ANTabla(INTabla,1)
          xtpri    = ANTabla(INTabla,2)
          xtfvig   = ANTabla(INTabla,3)
          xthvig   = ANTabla(INTabla,4)
          xtsadj   = ANTabla(INTabla,5)
          xtslnk   = ANTabla(INTabla,6)
          xfromusr = ANTabla(INTabla,7)
          xtmsg    = ANTabla(INTabla,8)


          select case xtipom
                 case 1
                      xtipomes = "GENERAL"
                 case 2
                      xtipomes = "PARTICULAR"
                 case 3
                      xtipomes = "PUBLIC"
                 case 4
                      xtipomes = "AUTORIZ"
                 case 5
                      xtipomes = "WORKFLOW"


          end select

          xfclas    = "dt"
          sImage    = "&nbsp;"

          select case xtpri
                 case 0

                      xfclas = "btbn"
                      sglmod = 5
                      sgltit = "Prioridad Alta."

                case 1

                     xfclas = "btba"
                     sglmod = 4
                     sgltit = "Prioridad Normal."

                case 2

                     xfclas = "btbs"
                     sglmod = 6
                     sgltit = "Prioridad baja."


          end select

          sglmod     = "<div class='" & xfclas & "' ><font class=btic >" & sglmod & "</font></div>"
          sImage     = ofv.GenLink("",xfclas & " title='" & sgltit & "' ","","","","",sglmod,"si")

          if  xtfvig = dFechai then
              XMARQ  = sglmar
          else
              XMARQ  = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
          end if

          IF  xtipom = 5 THEN
              xtfechavig = xthvig
          ELSE
              xtfechavig = fechaArg(cdate(xtfvig + clng(cdate("01-01-1900")))) & "&nbsp;&nbsp;" & xthvig
          END IF

          Xnlnk      = "<font color=#dddddd SIZE=1 >" & sgllnk & "</font>"
          Xnadj      = "<font color=#dddddd SIZE=1 >" & sgladj & "</font>"

          if xtsadj  = "S" then  Xnadj = sgladj
          if xtslnk  = "S" then  Xnlnk = sgllnk

                         Celdak = ""
                         Celdak = Celdak & ofv.GenCelda("","btbs","","","","","",sglprf,"si")
                         Celdak = Celdak & ofv.GenCelda("","ut","","","","","","&nbsp;<font color=#999999 ><u>" & xfromusr & "</u></font>&nbsp;&nbsp;","si")

                         Filak = ofv.GenRow("","","","","","","",Celdak,"si")
                         SArma = ofv.GenTabla("","aaa  ","","","","0","0","-1","-2","","",Filak,"si")


                         sHTML = sHTML &  "<tr >"
                         sHTML = sHTML &  "<td class=AAA colspan=8 style='line-height:4pt;' >&nbsp;</TD>"
                         sHTML = sHTML &  "</TR>"


                         sHTML = sHTML &  "<tr bgcolor=" & xTrColor & " >"
                         sHTML = sHTML &  "<td class=" & xfclas & " width=1% bgcolor=#FFFFFF >" & sImage & "</TD>"
                         sHTML = sHTML &  "<td class=vt width=55%  STYLE='text-align:justify;background-color:#DDDDDD;filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & Session("CMFPNTIT") & ", endColorStr=" & xTrColor & ", gradientType=1);' >&nbsp;&nbsp;<a class=tt >" & xtfechavig & "</a></TD>"
                         sHTML = sHTML &  "<td class=ut width=13% >&nbsp;</TD>"
                         sHTML = sHTML &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtipomes & "</font></TD>"
                         sHTML = sHTML &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtmcod & "</font></TD>"
                         sHTML = sHTML &  "<td class= width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnadj & "</TD>"
                         sHTML = sHTML &  "<td class= width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnlnk & "</TD>"
                         sHTML = sHTML &  "<td class=qt width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & XMARQ & "</TD>"
                         sHTML = sHTML &  "</TR>"

                         sHTML = sHTML &  "<tr bgcolor=" & xTrColor & " >"
                         sHTML = sHTML &  "<td style='cursor:pointer;border-style:solid;border-width:1;border-color:#eeeeee;' rowspan=2 width=1% class=ttr "
                     if  menu < 4 then
                         sHTML = sHTML &  "  title='Expirar mensaje " & xtmcod & "' > "

                         sHTML = sHTML &  "  <input class=bte type='button' value='X' onclick=""expmsg('" & xtmcod & "','" & xtfechavig & "');  "" >"
                     else
                        sHTML = sHTML &  "   >--- "
                     end if
                          sHTML = sHTML &  "</TD>"
                         sHTML = sHTML &  "<td class=uti colspan=7 STYLE='text-align:justify;' ><br><font color=#999999 >" & xtmsg & "</font></TD>"
                         sHTML = sHTML &  "</TR>"

                         sHTML = sHTML &  "<tr bgcolor=" & xTrColor & " height=20 >"
                         sHTML = sHTML &  "<td class=ut colspan=7 valign=top align=right style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' >" & sArma & "</TD>"
                         sHTML = sHTML &  "</TR>"


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF

      next


  end if

 '''form name='expira' action='noveexpira.asp?menu=" & menu & "&xmsg=' method='post' "

  sHTML = sHTML &  "</table>"
  sHTML = sHTML &  "</center>"

  sHTML = sHTML &  "<SCRIPT LANGUAGE='JAVASCRIPT' >"
  sHTML = sHTML &  " function expmsg(exmsg,exfec) {"
''''  sHTML = sHTML &  "  alert('A borrar: ' + exmsg);"
  sHTML = sHTML &  "  document.getElementById('MTEXP2A').innerHTML=exmsg;"
  sHTML = sHTML &  "  document.getElementById('MTEXP2B').innerHTML=exfec;"

  sHTML = sHTML &  "  document.getElementById('expira').action = document.getElementById('expira').action + exmsg;"
''''     sHTML = sHTML &  "  alert(document.getElementById('expira').action);"
  sHTML = sHTML &  "  document.getElementById('FDEXP').style.display='BLOCK';"
  sHTML = sHTML &  "}"
  sHTML = sHTML &  "</SCRIPT>"

 shtml = shtml & "<DIV id='FDEXP' STYLE='display:none;z-index:99;margin-top:1;margin-left:1;position:absolute;left:60px;top:100px;border-style:solid;border-width:2;border-color:" & Session("CMFPNTIT") & ";background-color:" & session("APLBGD") & ";"
 shtml = shtml & "'>"
 shtml = shtml & "<table id=""MTEXP1""  class=aaa width=400 cellspacing=0 cellpadding=2 "
 shtml = shtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"
 shtml = shtml & XspEx(Session("APLBGD"),"#BBBBBB",0) & ""
 shtml = shtml & chr(34)
 shtml = shtml & " >"
 shtml = shtml & "<tr>"
 shtml = shtml & "<td class=vt width=99% STYLE='color:" & session("CMFMNBOT") & ";filter:dropShadow(OFFY=1,OFFX=1,color=#f9f9f9);' >&nbsp;EXPIRACION DE MENSAJE</td>"
 shtml = shtml & "<td class=vt width=1% onclick=""document.getElementById('FDEXP').style.display='none';"" STYLE='color:" & session("CMFMNBOT") & ";filter:dropShadow(OFFY=1,OFFX=1,color=#f9f9f9);' >"
 shtml = shtml & "<b><font style='border-style:solid;border-width:1;border-color:#dddddd;' onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=corflaticon size=1 color=" & session("CMFMNBOT") & " >r</font></b>"
 shtml = shtml & "</td>"
 shtml = shtml & "</tr>"
 shtml = shtml & "</table>"

 ''sHTML = sHTML &  "<form name='expira' action='noveexpira.asp?menu=" & menu & "&xmsg=' method='post' >"
''sHTML = sHTML &  "<input type='submit' class=bt  value='CONFIRMAR'> "

 shtml = shtml & "<table id=""MTEXP2""  class=aaa width=400 cellspacing=4 cellpadding=2 "
 shtml = shtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
 shtml = shtml & " >"
 shtml = shtml & "<tr>"
 shtml = shtml & "<td class=tt width=25% >&nbsp;MENSAJE</td>"
 shtml = shtml & "<td class=vti width=75% id=""MTEXP2A""  >"
 shtml = shtml & Session("SegUsuario")
 shtml = shtml & "</td>"
 shtml = shtml & "</tr>"
 shtml = shtml & "<tr>"
 shtml = shtml & "<td class=tt   >&nbsp;FECHA</td>"
 shtml = shtml & "<td class=vti  id=""MTEXP2B""   >"
 shtml = shtml & session("cliente")
 shtml = shtml & "</td>"
 shtml = shtml & "</tr>"
 shtml = shtml & "<tr>"
 shtml = shtml & "<td class=ut colspan=2 align=center valign=bottom >"
 sHTML = sHTML &  "<form id='expira' action='noveexpira.asp?menu=" & menu & "&xmsg=' method='post' >"
 sHTML = sHTML &  "<input type='submit' class=bt  value='CONFIRMAR'></form>"
'' sHTML = sHTML &  "<input type='submit' class=bt  value='CONFIRMAR'>"
 shtml = shtml & "</td>"
 shtml = shtml & "</tr>"
 shtml = shtml & "</table>"
'' shtml = shtml & "</form>"
 shtml = shtml & "</DIV>"




  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML &  "</body></html>"



  response.write sHTML
end if

function EsqnomP(xtipo,xcodigo,ofv)

set xcn=ofv.conectar(ofv.strconn2)

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

call ofv.cerrarconn(xcn)

esqnomP = xdescri & ": " & xiden & " - " & xobjdes

end function


function fechaArg(f)
 dia=day(f)
 mes=month(f)
 ano=year(f)
 fechaarg=dia &"/" & mes & "/" & ano
end function

%>
