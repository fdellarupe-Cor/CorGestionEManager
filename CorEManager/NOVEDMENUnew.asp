<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
    response.expires=0

	'RESPONSE.Write(WKAPLIC & CHR(13) & CHR(13))

    wmnnoved = "NOVED" & WKAPLIC

    session(wmnnoved) = "SI"




    stilo1 = " style={background-color:" & session("CMFMENU") & ";border-style:outset;border-width:2;}"
    stilo2 = " style={background-color:" & session("CMFMNBOT") & ";}"


    ReDim Preserve bConsultar(10)

''''response.write "<br><br><br><br>" & Session("Form") & "<br>"

 Novedform = "CLIMN16317206150930885546227"

 if (session("UN") & "") = "" then


     XXXX = NEWEMPRE("NO")


     SEMPRE = ATREMP(0,0)
     session("UN") = SEMPRE & ""

     if (session("UN") & "") = "" then session("UN") = 1


 end if


    set cn = ofv.conectar(ofv.strconn2)
    set cn0 = ofv.conectar(ofv.strconn0)

wmnfuncion = ""

dim atit(50)
dim awmnasp(50)
wmntope = 10
wmntopeX = 10
WMNINITX = 2

 FOR TT=1 TO wmntope
     atit(tt) = "NO"
 NEXT

 IF WKAPLIC = "OR" OR WKAPLIC = "PL" THEN
    wmntopeX = 5
    WMNINITX = 3
 END IF


titc = 0
Strsql = "SELECT * FROM novedtit "
set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if not rsx.eof then

   for tt = WMNINITX to wmntopeX
       atit(tt) = rsx(tt)
       if WKAPLIC = "OR" then
          if tt = 3 then  atit(tt) = "ALERTAS"
       end if
       if rsx(tt) <> "NO" then
          OFV.ObtenerAtributosNOdo Novedform, "", "TIPO", CSTR(tt)
          bConsultar(TT) = OFV.AConsultar

          IF TT = 10 THEN

             IF bConsultar(TT) THEN
                session("SMPLMENU") = "N"
                atit(tt) = "NO"

             else
                if session("SMPLMENU") <> "S" then
                   session("SMPLMENU") = "S"
                   wmnfuncion = "<script language='javascript' > hdemenu(); </script> "
                end if
                titc = titc + 1



             end if


          ELSE

             IF bConsultar(TT) THEN
                 titc = titc + 1

             else

                 atit(tt) = "NO"

             end if

          END IF

       END IF
   next

   if  atit(10) <> "NO" then

       ii = 10
       xxxx = wmnbuscamenu("CLIS150172171122331051096140")
       wmntope = ii

   end if

end if
call ofv.cerrarconsulta(rsx)

menuSel = request.querystring("menu")
menu    = menuSel

if menu = "" then

   if   atit(3) <> "NO" then
        'menu = 3
        if WKAPLIC = "OR" then
           menu = 4
           menuSel = 4
        else
           menu=4
           menuSel = 4
        end if


   else
        menu = 0
        menuSel = 0

        for iix = 1 to 10
            if   atit(iix) <> "NO" then
                 menu = iix
                 menuSel = iix
                 exit for
            end if
        next
   end if

end if

menu    = cint(menu)

  if session("FormN") = "" then session("FormN") = "NOVEDADES"

  if session("SMPLMENU") = "S" then session("FormN") = "MENU PRINCIPAL"

  sHTML = ofv.FormHeader(session("FormN"))
'   shtml=shtml & "<link rel='stylesheet' href='windowfiles/dhtmlwindow.css' type='text/css' />"
'   shtml=shtml & "<script type=text/javascript src=windowfiles/dhtmlwindow.js>"

 '' sHTML = ofv.menuheader("","novedades ")

    xs = ""

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
      s = s & "menubar=no,toolbar=no,height=570,width=790,status=yes,resizable=yes" & Chr(34)
      s = s & " ); " & Chr(13) & "window.close(); " & Chr(13)
      s = s & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs & s & wmnfuncion
   ' shtml=shtml & "<link rel='stylesheet' href='windowfiles/dhtmlwindow.css' type='text/css' />"











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


   dim ANTabla(500,8)
   TNTabla = 0
   INTabla = 0
   xnohay1 = "S"
   xnohay2 = "S"
   SaleSel = "N"

Do until SaleSel = "S"

   select case menu

       case 1


       case 2
						'ESTE ES EL PUNTO DONDE ME INFORMA CUANDO ENTRA EN EMANAGER Y EL TITULO "INFORMATIVO"
                         session("ModMain") = "CL"

                         Strsql = "SELECT * FROM csnovedades "
                         Strsql = Strsql & " Where fecexp > " & dFechai
                         Strsql = Strsql & " and fecinicio < " & dFechai + 1
                         Strsql = Strsql & " And ("
                         Strsql = Strsql & " (tipo = 2 And usugrp = '" & session("usuario") & "') or "
                         Strsql = Strsql & " (tipo = 1 And usugrp = 'TODOS' and sector = 999 )    or "
                         Strsql = Strsql & " (tipo = 1 And usugrp = 'TODOS' and sector = 0" & session("UN") & ")"

           '              if session("SegGrpLst")="" then session("SegGrpLst")=Session("SegGrp")

               '          XMFTGRP2=split(session("SegGrpLst"),chr(9))

               '          XMFSEP  = ""
               '         GRPUSR=""
               '         for XMFI = 0 to ubound(XMFTGRP2)
                '         GRPUSR= GRPUSR & XMFSEP & "'GRP-" & XMFTGRP2(XMFI) &"'"

                '        XMFSEP=","
                '        next


                 '        Strsql = Strsql & " or ( usugrp in ("& GRPUSR &") )"

                         Strsql = Strsql & ") "
                         Strsql = Strsql & " order by fecinicio desc, tipo desc, importancia, objid desc"
						 'RESPONSE.Write(CHR(13) & CHR(13))
                         'response.write strsql
						 'RESPONSE.Write(CHR(13) & CHR(13))
                         set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         if   rsx.eof then
                              xnohay1 = "S"
                         else
                              xnohay1 = "N"
                           do until rsx.eof
                              INTabla = INTabla + 1

                              xfromusr = rsx(11)

                              xxxx = getInfoUser(rsx(11))

                              xfromusr = session("SGU_nombre")

                          '    strsql = "Select descripcion From Usuarios  where ID = '" & rsx(11) & "' "
                          '    set rsx1 = ofv.crearconsultaEx(StrSql,cn0,1,volver)
                          '    if  not rsx1.eof then
                          '        xfromusr = rsx1(0)
                          '    end if
                          '    call ofv.cerrarconsulta(rsx1)



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


                   IF WKAPLIC = "OR" THEN
					
					'ENTRA POR RIESGO OPERACIONAL EN LA SOLAPA DEL MENU "ALERTAS"
					
                       '''  RESPONSE.WRITE WKAPLIC & "<BR>"


                   strsql= "select         d.secuen, "
				   strsql =strsql & "     ISNULL(e.descri,D.DESCRIACC), g.descri, F.CODRIESGO, "
				   strsql =strsql & "a.* ,B2.comen1,B2.usr ,(select z.periodo from orperiodos z where z.objid=f.periodo) as periodo ,b1.fecvigen  "
				   strsql =strsql & "from csrelpathdescri a    "
				   strsql =strsql & "inner join objetos b on (a.cobj = b.codobjeto) "
				    strsql =strsql & "inner join objlnk c on (a.cobj = c.codobjeto) "
				   strsql =strsql & " inner join objaccion d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced) "
				   strsql =strsql & "LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)      "
				   strsql =strsql & "left join oractivriesgo f on (d.codaccion = f.actividad )     "
 				   strsql =strsql & "INNER join ORRSKCIRCUITO b1 on (f.tipobjeto=b1.tipobjeto and f.periodo=b1.periodo and  f.codobjeto = b1.codobjeto and  b1.estado=3)     "
				   strsql =strsql & "inner join OBJWFLDET b2 on (b1.objid = b2.codobjeto and b2.tipobjeto='RSK' AND B2.comen1='REFERENTE' and B2.usr='"&session("usuario")&"' )   "
				   strsql =strsql & "left join orimpriesgo g on "

                '   strsql =strsql & "(f.itemriesgo = g.objid) where " & clng(date())-1 &  "> fecvigen and d.critico = 'SI' and  F.requiereplan = 'SI' and   (F.incplan IS NULL OR F.incplan <> 'SI' )  order by a.tsec, D.secuen "
                    strsql =strsql & "(f.itemriesgo = g.objid) where " & dFechai - 29 &  "> fecvigen and d.critico = 'SI' and  F.requiereplan = 'SI' and   (F.incplan IS NULL OR F.incplan <> 'SI' )  order by a.tsec, D.secuen "
                   xnohay1 = "N"
				   'RESPONSE.Write(CHR(13) & CHR(13))
                   'RESPONSE.WRITE STRSQL
				   'RESPONSE.Write(CHR(13) & CHR(13))
					set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					if   rsx.eof then
						 xnohay1 = "S"

					else


                                   'xtmcod   = ANTabla(INTabla,0)
						           'xtipom   = ANTabla(INTabla,1)
						           'xtpri    = ANTabla(INTabla,2)
						           'xtfvig   = ANTabla(INTabla,3)
						           'xthvig   = ANTabla(INTabla,4)
						           'xtsadj   = ANTabla(INTabla,5)
						           'xtslnk   = ANTabla(INTabla,6)
						           'xfromusr = ANTabla(INTabla,7)
                                   'xtmsg    = ANTabla(INTabla,8)

						 do until rsx.eof

                              WKIPROCD = ""

							   FOR WKI=4 TO 11

								   WKIPROC = rsx(WKI) & ""

								   IF WKIPROC <> "" THEN
									  WKIPROCD = WKIPROCD & rsx(WKI) & " / "
								   END IF


							   NEXT

                               IF LEN(WKIPROCD) > 4 THEN
                                  WKIPROCD = LEFT(WKIPROCD,LEN(WKIPROCD)-3)
                               end if

                              INTabla = INTabla + 1
                              ANTabla(INTabla,0) = 0
                              ANTabla(INTabla,1) = 1
                              ANTabla(INTabla,2) = 0
                              ANTabla(INTabla,3) = 10
                              ANTabla(INTabla,4) = 10
                              ANTabla(INTabla,5) = "S"
                              ANTabla(INTabla,6) = "N"
                              ANTabla(INTabla,7) = RSX("usr")
                              ANTabla(INTabla,8) = "PLAN DE ACCION :<b>" & WKIPROCD &"</b><i>&nbsp;&nbsp;&nbsp;" & UCASE(RSX(1)&"<b></i> (" & rsx("periodo")&")</b>") &"&nbsp;&nbsp;&nbsp;" & GTFechadorSH(rsx("fecvigen"))
                              rsx.movenext

					      loop
					  end if
                        call ofv.cerrarconsulta(rsx)



                    '------------------------------------------------------
                     'indicadores vencidos

                   STRSQL= "select A.*,(select c.valor from orfrecuencia c where c.objid=a.unidadfrecuen) as dias from ORindicadores a,ORUSRLST b where a.objid=b.INDICADOR and b.USRRESP='USR-"& SESSION("USUARIO") &"' AND A.ESTADO=2"
				   'RESPONSE.Write(CHR(13) & CHR(13))
                   'response.write xnohay1
				   'RESPONSE.Write(CHR(13) & CHR(13))
                   set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                  ' if xnohay1 = "S" then xnohay1 = "N"
                   if   rsx.eof then
				      if  xnohay1<>"N" then	  xnohay1="S"
				   else
                   xnohay1="N"
				   do until rsx.eof

				     Svencido=validarFecuencia2(rsx(0),rsx("FECHAULTACT"),rsx("dias"))
				     if Svencido="SI" Then
                      INTabla=INTabla + 1
					  ANTabla(INTabla,0) = 0
					  ANTabla(INTabla,1) = 1
					  ANTabla(INTabla,2) = 0
					  ANTabla(INTabla,3) = 10
					  ANTabla(INTabla,4) = 10
					  ANTabla(INTabla,5) = "S"
					  ANTabla(INTabla,6) = "N"
					  ANTabla(INTabla,7) = ""
                      ANTabla(INTabla,8) = "INDICADOR VENCIDO : CODIGO :<I><B>"  & rsx(0) &"</B></I>&nbsp;----&nbsp;" & UCASE(RSX(1)&"")
                     end if
                     rsx.movenext

				  loop
				  end if
                        call ofv.cerrarconsulta(rsx)

                   ELSE
						'ESTE PUNTO ES PARA VER POR EMANAGER SOLAPA "DOCUMENTOS ACTIVOS"
                        session("ModMain") = "CL"

                        Strsql = "SELECT distinct * FROM csnovedades "
                        Strsql = Strsql & " Where fecexp > " & dFechai
                        Strsql = Strsql & " and fecinicio < " & dFechai + 1
                        Strsql = Strsql & " And tipo = 3"
                        Strsql = Strsql & " And ("
                        Strsql = Strsql & "    (usugrp = '" & session("usuario") & "') "
                        Strsql = Strsql & " or ( usugrp = 'TODOS' and sector = 999)"
                        Strsql = Strsql & " or ( usugrp = 'TODOS' and sector = 0" & session("UN") & ")"

                        Strsql = Strsql & "     ) "
                        Strsql = Strsql & " order by fecinicio desc, importancia, objid desc"
						'RESPONSE.Write(CHR(13) & CHR(13))
						'RESPONSE.Write(STRSQL)
						'RESPONSE.Write(CHR(13) & CHR(13))
                        set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                        if   rsx.eof then
                             xnohay1 = "S"

                        else

                             xvinoant = "no"
                             do until rsx.eof


                              if xvinoant <> rsx(1) then
                                xvinoant = rsx(1)


                               xvale = "N"

                               Strsql  = "SELECT * FROM objetos "
                               Strsql  = Strsql & " Where codobjeto = 0" & rsx(8)
                               set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                               if  not rsxp.eof then
                                   xvale = "S"
                               end if
                               call ofv.cerrarconsulta(rsxp)

                               IF  xvale = "S" THEN

                                xnohay1 = "N"
                                INTabla = INTabla + 1


                                xarc = "<font face=corflaticon color=#993333 size=4  >&#0157;</font>"


                                sActcon = "Consultar.asp?Codigo=" & rsx(8) & "&Tipo=" & rsx(9)
                                sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
                                sHref = "&nbsp;"
                              ' response.write sActcon
                               ' sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)
                               sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


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

                                sLink = OFV.GenLink("aaa title='Consultar Documento...' ", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")


                                xfromusr = rsx(11)


                                xxxx = getInfoUser(rsx(11))

                                xfromusr = session("SGU_nombre")


                            '    strsql   = "Select descripcion From Usuarios  where ID = '" & rsx(11) & "' "
                            '    set rsx1 = ofv.crearconsultaEx(StrSql,cn0,1,volver)
                            '    if  not rsx1.eof then
                            '        xfromusr = rsx1(0)
                            '    end if
                            '    call ofv.cerrarconsulta(rsx1)

                              ANTabla(INTabla,0) = rsx(1)
                              ANTabla(INTabla,1) = 3
                              ANTabla(INTabla,2) = clng(rsx(6))
                              ANTabla(INTabla,3) = clng(rsx(3))
                              ANTabla(INTabla,4) = rsx(12)
                              ANTabla(INTabla,5) = "S"
                              ANTabla(INTabla,6) = "N"
                              ANTabla(INTabla,7) = xfromusr
                              ANTabla(INTabla,8) = sLink

                             end if
                            end if

                            rsx.movenext
                            loop
                        end if
                        call ofv.cerrarconsulta(rsx)

                     END IF

        case 4

                        session("ModMain") = "WK"
                        Session("Segce")   = "S"
                        Session("Segco")   = "S"

						'SOLAPA "AUTORIZACION DE DOCUMENTOS"
                        Strsql = "SELECT * FROM csnovedades "
                        Strsql = Strsql & " Where fecexp > " & dFechai
                        Strsql = Strsql & " and fecinicio < " & dFechai + 1
                        Strsql = Strsql & " And tipo = 4"
                        Strsql = Strsql & " And usugrp = '" & session("usuario") & "' order by fecinicio desc, importancia, objid desc"
						'RESPONSE.Write(CHR(13) & CHR(13))
                        'response.write strsql
						'RESPONSE.Write(CHR(13) & CHR(13))
                        set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                        if  rsx.eof then
                            xnohay1 = "S"
                        else
                            xnohay1 = "N"
                            do  until rsx.eof


                            INTabla = INTabla + 1

                            Strsql = "SELECT archivo FROM tipobjeto "
                            Strsql = Strsql & " Where tipobjeto = '" & rsx(9) & "'"
							'RESPONSE.Write(CHR(13) & CHR(13))
	                        'response.write strsql
							'RESPONSE.Write(CHR(13) & CHR(13))
                            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                            if  not rs.eof then
                                xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
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


                            xxxx = getInfoUser(rsx(11))
                            xfromusr = session("SGU_nombre")


                          '  strsql   = "Select descripcion From Usuarios  where ID = '" & rsx(11) & "' "
                          '  set rsx1 = ofv.crearconsultaEx(StrSql,cn0,1,volver)
                          '  if  not rsx1.eof then
                          '      xfromusr = rsx1(0)
                          '  end if
                          '  call ofv.cerrarconsulta(rsx1)


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

                         xtipos = " OBJTIPO in ('RSK'"
                         separ = ","
                         Strsql = "SELECT tipobjeto FROM tipobjeto "
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
                         xtipos = xtipos & ")"

                         SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & Session("Usuario") & "') OR ("

                         xxxx     = getInfoUser(Session("Usuario"))
                         xusrlog  = session("SGU_nombre")



                         if  xusrlog <> "NO" THEN


                             SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "') ) "
                             SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

'''''''''''''''''''''''''''''''''''''''''''
                             if  Session("DELLOG") = "S" and session("SegGrpLst")<> "" then

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
						'RESPONSE.Write(CHR(13) & CHR(13))
                        'response.write strsql
						'RESPONSE.Write(CHR(13) & CHR(13))

                         set rsx = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
                         if  rsx.eof then
                             xnohay2 = "S"
                         else

                             do  until rsx.eof

                               xvale = "N"

								'POR ESTE IF SE DECIDE SI ES RIESGO OPERACIONAL O EMANAGER, SI ES OR = RIESGO OPERACIONAL SI ES EM = EMANAGER
								'RESPONSE.Write(CHR(13) & CHR(13))
		                        'response.write WKAPLIC
								'RESPONSE.Write(CHR(13) & CHR(13))
								IF WKAPLIC = "OR" THEN
									if rsx(3) = "RSK" then
										Strsql  = "SELECT * FROM ORRSKCIRCUITO "
										Strsql  = Strsql & " Where OBJID = 0" & rsx(4)
										set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
										if  not rsxp.eof then
										   xvale = "S"
										   xMAINAPL = rsxp("APLIC")
										end if
										call ofv.cerrarconsulta(rsxp)
									END IF
							   ELSE
							   		if rsx(3) <> "RSK" then		
									'RESPONSE.Write("ENTRO POR ACA TAMBIEN")
									Strsql  = "SELECT * FROM objetos "
									Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
									set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
									if  not rsxp.eof then
										xvale = "S"
										xMAINAPL = rsxp("APLIC")
									end if
									call ofv.cerrarconsulta(rsxp)
									END IF
								END IF

                               IF  xvale = "S" THEN

                                 xnohay2 = "N"

                                 INTabla = INTabla + 1


                                 xarc = "<font face=corflaticon color=#005500 size=4  >&#0157;</font>"

                                 Session("ASP")      = "novedmenu.asp?menu=4"
                                 session("sEstado")  = "PARA APROBAR"
                                 session("sEstadon") = "2"
                                 Session("Form2")    = "apr"

                                 aspeditar = "Editar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & RSX(8) & "&MAINAPL=" & xMAINAPL


                               ''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

                                 aspeditar = cstr(aspeditar)
                                 sHref     = "&nbsp;"
                        '        response.write sActcon
                                 sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                                 wNomL     = "&nbsp;CIRCUITO:&nbsp;" & RSX(7) & "&nbsp;&nbsp;&nbsp;PASO:&nbsp;"  & RSX(16) & "&nbsp;"

                                 sImagen   = wNomL

                                 if xarc <> "NO" then
                                    sImagen = xarc & "&nbsp;&nbsp;" & EsqnomP(rsx(3),rsx(4),ofv) & "&nbsp;&nbsp;" & sImagen
                                 end if

                                 sLink     = OFV.GenLink("aaa title='Autorizar documento ...' ", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")

                                 xfromusr = rsx(6)


                                 xxxx = getInfoUser(rsx(6))
                                 xfromusr = session("SGU_nombre")

                                 ANTabla(INTabla,0) = "CKT" & INTabla
                                 ANTabla(INTabla,1) = 5
                                 ANTabla(INTabla,2) = 0
                                 ANTabla(INTabla,3) = dFechai
                                 ANTabla(INTabla,4) = rsx(10)
                                 ANTabla(INTabla,5) = "S"
                                 ANTabla(INTabla,6) = "N"
                                 ANTabla(INTabla,7) = xfromusr
                                 ANTabla(INTabla,8) = sLink

                               end if



                             rsx.movenext
                             loop
                        end if
                        call ofv.cerrarconsulta(rsx)

       case 5
					
					
					'ESTA ES LA OPCION DE ULTIMOS CONSULTADOS DENTRO DE LAS NOVEDADES
                        session("ModMain") = "CL"

    cuenta = 0

''''''''''''''''''''''''''''''''
    XWHERE = " and USUARIO = '" & session("usuario") & "' "
    XWHERE = XWHERE & " AND CODOBJETO > 0 "

   if wkaplic = "OR" or wkaplic = "PL" then

      XWHERE = XWHERE & " AND TIPOBJETO = 'RSK' "

   else

      XWHERE = XWHERE & " AND TIPOBJETO <> 'RSK' "

   end if

    XORDER = " group by codobjeto ORDER BY fecha desc "


                 strsql = ""
                 strsql = strsql & "SELECT CODOBJETO, max(TIPOBJETO), max(finicio) as fecha  "
                 strsql = strsql & " FROM OBJALOGDET B1"
                 strsql = strsql & " WHERE TACC = 'CONSEN'  " & XWHERE

                 strsql = strsql & XORDER

				'RESPONSE.Write(CHR(13) & CHR(13))
				'RESPONSE.Write(STRSQL)
				'RESPONSE.Write(CHR(13) & CHR(13))

                        set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                        if   rsx.eof then
                             xnohay1 = "S"

                        else

                             xvinoant = "no"
                             do until rsx.eof


                               xvale = "N"



                               if wkaplic = "OR" or wkaplic = "PL" then


                                   Strsql  = "SELECT * FROM ORRSKCIRCUITO "
                                   Strsql  = Strsql & " Where OBJID = 0" & rsx(0)
                                   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                   if  not rsxp.eof then
                                       xvale = "S"
                                       xMAINAPL = rsxp("APLIC")
                                   end if
                                   call ofv.cerrarconsulta(rsxp)

                                   WKSPTIPOOBJ = "RSK"


                               else

                                   Strsql  = "SELECT * FROM objetos "
                                   Strsql  = Strsql & " Where codobjeto = 0" & rsx(0)
                 ''    RESPONSE.WRITE STRSQL & "<BR>"
                                   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                   if  not rsxp.eof then
                                       xvale = "S"
                                       xMAINAPL = rsxp("APLIC")
                                   end if
                                   call ofv.cerrarconsulta(rsxp)


                                   WKSPTIPOOBJ = "VER"

                               end if




                               IF  xvale = "S" THEN


                                cuenta = cuenta + 1

                                xnohay1 = "N"
                                INTabla = INTabla + 1


                                xarc = "<font face=corflaticon color=#993333 size=4  >&#0157;</font>"


                                sActcon = "Consultar.asp?Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
                                sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
                                sHref = "&nbsp;"
                              ' response.write sActcon
                             ''   sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

                                sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


                                sTipobjeto = EsqnomP(WKSPTIPOOBJ,RSX(0),ofv)






                                wmsgx = sTipobjeto & ""
                                wNomL = "&nbsp;" & wmsgx & "&nbsp;"



                                sImagen = wNomL

                                if xarc <> "NO" then
                                   sImagen = xarc & "&nbsp;" & sImagen
                                end if

                                sLink = OFV.GenLink("aaa title='Consultar Documento...' ", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")


                                xfromusr = "COREMANAGER"


                                xxxx = getInfoUser(xfromusr)

                                xfromusr = session("SGU_nombre")


                              ANTabla(INTabla,0) = "0"
                              ANTabla(INTabla,1) = 3
                              ANTabla(INTabla,2) = 1
                              ANTabla(INTabla,3) = dFechai
                              ANTabla(INTabla,4) = 0
                              ANTabla(INTabla,5) = "S"
                              ANTabla(INTabla,6) = "N"
                              ANTabla(INTabla,7) = xfromusr
                              ANTabla(INTabla,8) = sLink

                             end if

                            if cuenta > 9 then exit do


                            rsx.movenext
                            loop
                        end if
                        call ofv.cerrarconsulta(rsx)



        case 6

                        session("ModMain") = "WK"
                        Session("Segce")   = "S"
                        Session("Segco")   = "S"



                        set cn1 = ofv.conectar(ofv.strconn5)
                        set cn4 = ofv.conectar(ofv.strconn6)

                         xtipos = " OBJTIPO in ("
                         separ = ""
                         Strsql = "SELECT tipobjeto FROM tipobjeto "
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
                         xtipos = xtipos & ")"

                         SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & Session("Usuario") & "') OR ("

                         xxxx     = getInfoUser(Session("Usuario"))
                         xusrlog  = session("SGU_nombre")




                         if  xusrlog <> "NO" THEN


                             SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "') ) "
                             SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

'''''''''''''''''''''''''''''''''''''''''''
                             if  Session("DELLOG") = "S" and session("SegGrpLst")<> "" then

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

                      ''' RESPONSE.WRITE STRSQL & "<BR>"
                         set rsx = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
                         if  rsx.eof then
                             xnohay2 = "S"
                         else

                             do  until rsx.eof

                               xvale = "N"

                               Strsql  = "SELECT * FROM objetos "
                               Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
                               set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                               if  not rsxp.eof then
                                   xvale = "S"
                               end if
                               call ofv.cerrarconsulta(rsxp)

                               IF  xvale = "S" THEN

                                 xnohay2 = "N"

                                 INTabla = INTabla + 1

                                 Strsql  = "SELECT archivo FROM tipobjeto "
                                 Strsql  = Strsql & " Where tipobjeto = '" & rsx(3) & "'"
                                 set rs  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                 if  not rs.eof then
                            '         xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
                           '      else
                           '          xarc = "NO"
                                 end if
                                 call ofv.cerrarconsulta(rs)

                                 xarc = "<font face=corflaticon color=#005500 size=4  >&#0157;</font>"

                                 Session("ASP")      = "novedmenu.asp?menu=4"
                                 session("sEstado")  = "PARA APROBAR"
                                 session("sEstadon") = "2"
                                 Session("Form2")    = "apr"

                                 aspeditar = "Editar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & RSX(8)


                               ''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

                                 aspeditar = cstr(aspeditar)
                                 sHref     = "&nbsp;"
                        '        response.write sActcon
                                 sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                                 wNomL     = "&nbsp;CIRCUITO:&nbsp;" & RSX(7) & "&nbsp;&nbsp;&nbsp;PASO:&nbsp;"  & RSX(16) & "&nbsp;"

                                 sImagen   = wNomL

                                 if xarc <> "NO" then
                                    sImagen = xarc & "&nbsp;&nbsp;" & EsqnomP(rsx(3),rsx(4),ofv) & "&nbsp;&nbsp;" & sImagen
                                 end if

                                 sLink     = OFV.GenLink("aaa title='Autorizar documento ...' ", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")

                                 xfromusr = rsx(6)


                                 xxxx = getInfoUser(rsx(6))
                                 xfromusr = session("SGU_nombre")


                                 ANTabla(INTabla,0) = "0"
                                '' ANTabla(INTabla,0) = "PEN" & INTabla
                                 ANTabla(INTabla,1) = 5
                                 ANTabla(INTabla,2) = 0
                                 ANTabla(INTabla,3) = dFechai
                                 ANTabla(INTabla,4) = rsx(10)
                                 ANTabla(INTabla,5) = "S"
                                 ANTabla(INTabla,6) = "N"
                                 ANTabla(INTabla,7) = xfromusr
                                 ANTabla(INTabla,8) = sLink

                               end if



                             rsx.movenext
                             loop
                        end if
                        call ofv.cerrarconsulta(rsx)





       case 10

                        session("ModMain") = "CL"

                                xnohay1 = "N"


                                sLink = "<INPUT TYPE=BUTTON VALUE='HIDE MENU' onclick='hdemenu();' ><INPUT TYPE=BUTTON VALUE='SHOW MENU' onclick='mostrarmenu();' >"


                                 INTabla = INTabla + 1


                              ANTabla(INTabla,0) = "0"
                              ANTabla(INTabla,1) = 3
                              ANTabla(INTabla,2) = 1
                              ANTabla(INTabla,3) = dFechai
                              ANTabla(INTabla,4) = 0
                              ANTabla(INTabla,5) = "S"
                              ANTabla(INTabla,6) = "N"
                              ANTabla(INTabla,7) = "COREMANAGER"
                              ANTabla(INTabla,8) = sLink












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


''''''''''''''''''''''''''''
     sHTML   = sHTML &  "<table  width=100% cellspacing=0 cellpadding=2 class='HX1CL'  style='border-style:solid;border-width:1;border-color:#ffffff;"

' IF  Session("FULLGUI") = "N" THEN
     shtml = shtml & "background-color:#cccccc;"
' else
'     shtml = shtml & XspEx("#eeeeee",Session("CMFPNTIT"),0) & ";"
' end if

 shtml = shtml & "' >"

 xnewmenu = ""
 xoldmenu = ""

     sHTML   = sHTML &  "<tr >"

     sHTML   = sHTML &  "<td class='UT'  >"

             sHTML = sHTML &  "<table  cellspacing=1 cellpadding=3  >"

             for ii=2 to wmntope
                 if  atit(ii) <> "NO" then

                    if  ii = 10 then
                         stilo  = stilo2
                    else

                        if  ii = menu  then
                            xsclase = "ttr STYLE='border-style:inset;border-color:" & Session("CMFPNTITFT") & ";border-width:1;padding-left:5px;padding-right:5px;" & XspEx("#f7aa22","#dddddd",0) & ";' "
                        else
                            xsclase = "ttr STYLE='border-right-style:solid;border-right-color:" & Session("CMFPNTITFT") & ";border-right-width:1;padding-left:5px;padding-right:5px;' "
                        end if

                        xmDato = REPLACE(atit(ii)," ","&nbsp;")

                        if ii < 11 then
                           xmhref    = "<a class='xts' href='NOVEDMENU.asp?menu=" & ii & "' >" & xmDato & "</a>"
                           xoldmenu  = xoldmenu &  "<td class=" & xsclase & " align=center  >"
                           xoldmenu  = xoldmenu &  xmhref & "</TD>"
                        else
                           xmhref    = "<a class='xts' href='" & awmnasp(ii) & "' >" & xmDato & "</a>"
                           xnewmenu  = xnewmenu &  "<td class=" & xsclase & " align=center  >"
                           xnewmenu  = xnewmenu &  xmhref & "</TD>"

                        end if



                    end if

                 end if
            next

            sHTML  = sHTML &  xnewmenu & xoldmenu

            sHTML = sHTML &  "</table>"

  sHTML = sHTML &  "</TD>"
  sHTML = sHTML &  "</TR>"
  sHTML = sHTML &  "</table>"


'''''''''''''''''''''''''''''



  sHTML = sHTML &  "<center>"

 shtml = shtml & "<div id='WMN0' style='overflow:auto;height:50%;" & WStyl & "' >"


  if  xnohay1 = "S" and xnohay2 = "S" then

      sHTML = sHTML &  "<table  margin=3 width=95% cellspacing=0 cellpadding=0  style='border-style:solid;border-width:1;border-color:#eeeeee;'  >"
      sHTML = sHTML &  "<tr  height=200 >"
      sHTML = sHTML &  "<td class=ttt width=100% align=center valign=middle ><big>SIN&nbsp;NOVEDADES&nbsp;DE&nbsp;" & atit(menu) & "&nbsp;</TD>"
      sHTML = sHTML &  "</TR>"

  else
      sHTML = sHTML &  "<table bgcolor=#eeeeee margin=3 width=95% cellspacing=0 cellpadding=0  style='border-style:solid;border-width:1;border-color:#eeeeee;'  >"

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

                     if xtmcod <> "0" then
                         sHTML = sHTML &  "<tr >"
                         sHTML = sHTML &  "<td class=AAA colspan=8 style='line-height:2pt;' >&nbsp;</TD>"
                         sHTML = sHTML &  "</TR>"



                         sHTML = sHTML &  "<tr bgcolor=" & xTrColor & " >"
                         sHTML = sHTML &  "<td class=" & xfclas & " width=1% bgcolor=#FFFFFF >" & sImage & "</TD>"
                         sHTML = sHTML &  "<td class=vt width=55%  STYLE='text-align:justify;background-color:#DDDDDD;filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & Session("CMFPNTIT") & ", endColorStr=" & xTrColor & ", gradientType=1);' >&nbsp;&nbsp;<a class=tt >" & xtfechavig & "</a></TD>"
                         sHTML = sHTML &  "<td class=ut width=13% >&nbsp;</TD>"


                       if session("NOVFULL") = "S" THEN

                         sHTML = sHTML &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtipomes & "</font></TD>"
                         sHTML = sHTML &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtmcod & "</font></TD>"
                         sHTML = sHTML &  "<td class=btba width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnadj & "</TD>"
                         sHTML = sHTML &  "<td class=btbn width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnlnk & "</TD>"
                         sHTML = sHTML &  "<td class=qt width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & XMARQ & "</TD>"
                       ELSE
                         sHTML = sHTML &  "<td class=UT COLSPAN=3 style='border-style:solid;border-width:1;border-color:#eeeeee;' >&nbsp;</TD>"
                         sHTML = sHTML &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtipomes & "</font></TD>"
                         sHTML = sHTML &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtmcod & "</font></TD>"
                       END IF


                         sHTML = sHTML &  "</TR>"

                    end if

                         sHTML = sHTML &  "<tr bgcolor=" & xTrColor & " >"
                         sHTML = sHTML &  "<td class=uti colspan=8 STYLE='text-align:justify;line-height:30px;' ><font color=#999999 >" & xtmsg & "</font></TD>"
                         sHTML = sHTML &  "</TR>"



                     if session("NOVFULL") = "S" THEN

                       if xtmcod <> "0" then


                         sHTML = sHTML &  "<tr bgcolor=" & xTrColor & " height=20 >"
                         sHTML = sHTML &  "<td class=ut colspan=8 valign=top align=right style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' >" & sArma & "</TD>"
                         sHTML = sHTML &  "</TR>"

                       end if

                     ELSE

                       if xtmcod <> "0" then


                         sHTML = sHTML &  "<tr bgcolor=" & xTrColor & " height=20 >"
                         sHTML = sHTML &  "<td class=ut colspan=8 valign=top align=right  >&nbsp;</TD>"
                         sHTML = sHTML &  "</TR>"

                       end if

                     END IF


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF

      next


  end if



  sHTML = sHTML &  "</table>" '& session("LUID") & "---" &   session("LPWD")


 shtml = shtml & "</div>"

  sHTML = sHTML &  "</center>"

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

'  shtml=shtml &"<script type=text/javascript>"

' shtml=shtml  &  " var googlewin=dhtmlwindow.open('googlebox', 'iframe', 'http://images.google.com/', '#1: Google Web site', 'width=30px,height=350px,resize=1,scrolling=1,center=1', 'recal');"
' shtml=shtml  &  " var googlewin1=dhtmlwindow.open('googlebox1', 'iframe', 'http://images.google.com/', '#1: Google Web site', 'width=30px,height=350px,resize=1,scrolling=1,center=1', 'recal');"
' shtml=shtml  &  " var googlewin2=dhtmlwindow.open('googlebox2', 'iframe', 'http://images.google.com/', '#1: Google Web site', 'width=30px,height=350px,resize=1,scrolling=1,center=1', 'recal');"

 '" var inlinewin=dhtmlwindow.open('broadcastbox', 'inline', 'This is some inline content. <a href='http://dynamicdrive.com'>Dynamic Drive</a>', '#2: Broadcast Title', 'width=300px,height=120px,left=150px,top=10px,resize=1,scrolling=0', 'recal')"

'shtml=shtml & "</script>"


  sHTML = sHTML &  "</body></html>"



  response.write sHTML
end if

function EsqnomP(xtipo,xcodigo,ofv)

         set xcn = ofv.conectar(ofv.strconn2)


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

    sTipobjetoX  = xdescri & ": "  & xobjdes

 ELSE


         sTipobjetoX  = xcodigo

         Strsql2 = "Select * From CSOBJETOSRDO Where CODOBJETO = 0" & xcodigo
         set rsM = ofv.crearconsultaEx(StrSql2,xcn,1,parametros)
         IF NOT RSM.EOF THEN

           sTipobjeto1 = RSM(1) & ":&nbsp;"
           sTipobjeto2 = RSM(3)
           sTipobjeto3 = " - " & RSM(2)
           sTipobjeto4 = " ("  & RSM(5) & " "
           sTipobjeto5 = RSM(6) & ") "

           if session("FULLDOCN") = "1" then
              sTipobjetoX  = sTipobjeto1 & sTipobjeto2
           elseif session("FULLDOCN") = "2" then
              sTipobjetoX  = sTipobjeto1 & sTipobjeto2 & sTipobjeto3
           else
              sTipobjetoX  = sTipobjeto1 & sTipobjeto2 & sTipobjeto3 & sTipobjeto4 & sTipobjeto5
           end if



         END IF

 END IF





         call ofv.cerrarconn(xcn)


         esqnomP = sTipobjetoX


end function


function fechaArg(f)
 dia=day(f)
 mes=month(f)
 ano=year(f)
 fechaarg=dia &"/" & mes & "/" & ano
end function

function wmnbuscamenu(wmnSubmenu)

          Set cx0 = OFV.conectar(OFV.strconn0)
          Set cx1 = OFV.conectar(OFV.strconn1)

            xsel = 1
          StrSql = "Select * From Menux Where mdepen='" & wmnSubmenu & "' and mcodapli = '" & Session("CorApli") & "' "

          Set rs = OFV.crearconsultaEx(StrSql, cx1, 1, "")
            If Not rs.EOF Then
               If rs.fields.Count > 6 Then
                  xsel = 6
                  Call OFV.cerrarconsulta(rs)
                  StrSql = "Select * From Menux Where mdepen='" & wmnSubmenu & "' and mcodapli = '" & Session("CorApli") & "' order by Morden "

                  Set rs = OFV.crearconsultaEx(StrSql, cx1, 1, "")
               Else
                  xsel = 1
               End If
            End If
          Do Until rs.EOF



              atrok = "N"
              StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'M' "
              StrSql = StrSql & " And alias = '" & rs(xsel) & "'  "
              StrSql = StrSql & " And usuario = '" & Session("segperfil") & "' And "
              StrSql = StrSql & "  Aplicacion = '" & Session("CorApli") & "'"

              Set rs1 = OFV.crearconsultaEx(StrSql, cx0, 1, "")
              If rs1.EOF Then
                 atrok = "O"
              ElseIf rs1(1) Then
                     atrok = "S"
                  Else
                     atrok = "N"
              End If
              Call OFV.cerrarconsulta(rs1)

              xcount  = 0

              If atrok = "O" Then

                 StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'M' "
                 StrSql = StrSql & " And alias = '" & rs(xsel) & "' And  "

                 if  Session("DELLOG") = "S" and session("SegGrpLst") <> "" then

                     XMFTGRP = split(session("SegGrpLst"),chr(9))
                     xcount  = ubound(XMFTGRP) + 1
                     XMFSEP  = ""
                     StrSql = StrSql & " grupo in ("

                     for XMFI = 0 to ubound(XMFTGRP)
                         StrSql  = StrSql  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                         XMFSEP  = ","
                     next

                     StrSql = StrSql & ") "

                 else
                     StrSql = StrSql & " grupo = '" & Session("seggrp") & "' "
                 end if


                 StrSql = StrSql & " And Aplicacion = '" & Session("CorApli") & "'"
                 Set rs2 = OFV.crearconsultaEx(StrSql, cx0, 1, "")
                 If  rs2.EOF Then
                        atrok = "O"
                 Else
                     if  Session("DELLOG") = "S" then
                         XMFSUM = 0
                         XMFSUMN = 0
                         do until rs2.eof
                            XMFSUM = XMFSUM + 1
                            If rs2(1) Then
                               atrok = "S"
                               exit do
                            Else
                               atrok   = "N"
                               XMFSUMN = XMFSUMN + 1
                            End If
                         rs2.movenext
                         loop

                         if  atrok = "N" then
                             if  xcount >  XMFSUMN then atrok = "S"
                         end if

                     else
                         If rs2(1) Then
                            atrok = "S"
                         Else
                            atrok = "N"
                         End If
                     end if
                 End If
                 Call OFV.cerrarconsulta(rs2)

              End If


              If atrok = "O" Or atrok = "S" Then


                          sInput = Trim(rs(xsel))
                          sHref = Trim(rs(2))
                              sTarget = "info"
                              sPath = Session("CorPath")
                              If rs(3) <> "" Then atrok = "N"
                              spt = InStr(rs(2), ".asp")
                              xesm = "N"
                              If spt > 1 Then
                                  If Left(rs(2), spt - 1) = "Submenus" Then atrok = "N"
                              End If
                              sseg = InStr(sHref, "?")
                              If sseg = 0 Then
                                  sHref = sHref & "?"
                              Else
                                  sHref = sHref & "&"
                              End If
                              sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rs(1)), " ", "_")

                              If atrok <> "N" then

                                 ii = ii + 1

                                 atit(ii)    = UCase(rs(1))
                                 awmnasp(ii) = sPath & sHref

                              end if





                  End If




              rs.movenext
          Loop
          Call OFV.cerrarconsulta(rs)
          Call OFV.cerrarconn(cx1)
          Call OFV.cerrarconn(cx0)




end function


function validarFecuencia2(Idindicador,fechaultact,dias)


  sinput="SI"

  FEC=cdate( dias + CLNG("0" & fechaultact) + clng(cdate("01-01-1900") ))

  if DATE()>= fec then
    SInput="SI"
   ELSE
    SINPUT="NO"
   end if
  validarFecuencia2=sinput '& fec & "--" & date()
end function



%>
