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
    session("novedgp") = "SI"
    stilo1 = " style={background-color:" & session("CMFMENU") & ";border-style:outset;border-width:2;}"
    stilo2 = " style={background-color:" & session("CMFMNBOT") & ";}"

 OBJTIPO = request.querystring("OBJTIPO")
 if OBJTIPO = "" then
    if session("OBJTIPO") <> "" then
       OBJTIPO = session("OBJTIPO")
     else
       OBJTIPO = "REQ"
       session("OBJTIPO") = OBJTIPO

     end if
 else
    session("OBJTIPO") = OBJTIPO
 end if 


    ReDim Preserve bConsultar(10)

 Novedform = "CLIMN16317206150930885546227"


    set cn = ofv.conectar(ofv.strconn6)
    set cn0 = ofv.conectar(ofv.strconn0)

wmnfuncion = ""

dim atit(50)
dim awmnasp(50)
wmntope = 10
titc = 0
Strsql = "SELECT * FROM novedtit "
set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if not rsx.eof then

   for tt = 2 to 10
       atit(tt) = rsx(tt)
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

   if   atit(2) <> "NO" then
        menu = 2
        menuSel = 2
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
      s = s & "resizable=yes,menubar=no,toolbar=no,height=570,width=790" & Chr(34)
      s = s & " ); " & Chr(13) & "window.close(); " & Chr(13)
      s = s & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs & s & wmnfuncion









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

            session("ModMain") = "WK"



                        SEGWHERE = " WHERE OBJTIPO = '" & session("OBJTIPO") & "' AND  PSCFIN = 'NO' AND ( (RESPASG = '" & Session("Usuario") & "') OR ("

                         xxxx     = getInfoUser(Session("Usuario"))
                         xusrlog  = session("SGU_nombre")


                         if  xusrlog <> "NO" THEN


                             SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "') ) "
                             SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "


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





                             if  cstr(session("SGU_idcostcenter") & "") <> "0" then
                                 SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
                                 SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & session("SGU_idcostcenter") & "' OR SUBCODIGO = '" & session("SGU_idcostcenter") & "' ) "
                                 SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
                                 SEGWHERE = SEGWHERE & " ) ) "
                             END IF


                         END IF

                         SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "

                         Strsql = "SELECT * FROM CKTVIEW " & SEGWHERE

                      '''   RESPONSE.WRITE STRSQL & "<BR>" & chr(13)
                         set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         if  rsx.eof then
                             xnohay2 = "S"
                         else

                             do  until rsx.eof

                               xvale = "N"


                               StrSql = "Select * From REQUERIMIENTO where codigo = '" & RSX(4) & "' "
                               set rsxp = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                               if NOT rsxp.eof then
                                  xCLI = rsxp(2) 
                                  xCLIN = rsxp(0) 
                                  xvale = "S"
                               end if

                               call ofv.cerrarconsulta(rsxp)    
    
                               IF  xvale = "S" THEN 

                                 xnohay2 = "N"

                                 INTabla = INTabla + 1


                                 xarc = "<font face=webdings color=#993333 size=4  >&#0157;</font>"

                                 sHref     = "&nbsp;"
                                 sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('editarPRY.asp?codigo=" &  RSX(4) & "&tipo=REQ&SEGCKTPSC=" & RSX(8) & "'); return false;" & Chr(34)

                                 wNomL = "&nbsp;REQUERIMIENTO: <b>"  & xCLIN & "-" & XCLI & "</b>&nbsp;&nbsp;&nbsp;CIRCUITO: <b>" & RSX(7) & "</b>&nbsp;&nbsp;&nbsp;PASO: <b>"  & RSX(16) & "</b>"

                                 sImagen   = wNomL

                                 if xarc <> "NO" then
                                    sImagen = xarc & sImagen
                                 end if

                                 sLink     = OFV.GenLink("aaa title='Autorizar circuito ...' ", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")

                                 xfromusr = rsx(6)


                                 xxxx = getInfoUser(rsx(6))
                                 xfromusr = session("SGU_nombre")



                                 ANTabla(INTabla,0) = "CKT" & INTabla
                                 ANTabla(INTabla,1) = 5
                                 ANTabla(INTabla,2) = 0
                                 ANTabla(INTabla,3) = rsx(10)
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


   TNTabla = 0
   INTabla = 0


'''''''''''''' version 2


                       SEGWHERE = " WHERE OBJTIPO = '" & session("OBJTIPO") & "' AND  SEGESTADO = 'NO' AND ( (RESPASG = '" & Session("Usuario") & "') OR ("

                         xxxx     = getInfoUser(Session("Usuario"))
                         xusrlog  = session("SGU_nombre")


                         if  xusrlog <> "NO" THEN


                             SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PASORSP = '" & Session("Usuario") & "') ) "
                             SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "


                             if  Session("DELLOG") = "S" then

                                 XMFTGRP = split(session("SegGrpLst"),chr(9))
                                 xcount  = ubound(XMFTGRP) + 1
                                 XMFSEP  = "" 
                                 SEGWHERE = SEGWHERE & " PASORSP in ("

                                 for XMFI = 0 to ubound(XMFTGRP)
                                     SEGWHERE = SEGWHERE  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                                     XMFSEP  = "," 
                                 next

                                 SEGWHERE = SEGWHERE & ") ) )"

                            else
                                 SEGWHERE = SEGWHERE & " PASORSP = '" & SESSION("SEGGRP") & "') ) "
                            end if 





                             if  cstr(session("SGU_idcostcenter") & "") <> "0" then
                                 SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
                                 SEGWHERE = SEGWHERE & " (PASORSP = '" & session("SGU_idcostcenter") & "') "
                                 SEGWHERE = SEGWHERE & " AND (NOT SEGNIVELCNF < " & SESSION("SEGNIVEL") & " ) "
                                 SEGWHERE = SEGWHERE & " ) ) "
                             END IF


                         END IF

                         SEGWHERE = SEGWHERE & " ) ) ORDER BY SEGFINICIO desc,SEGHINICIO desc "

                         Strsql = "SELECT * FROM CKTVIEWnv " & SEGWHERE

                     '''    RESPONSE.WRITE STRSQL & "<BR>" & chr(13)
                         set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         if  rsx.eof then
                             xnohay2 = "S"
                         else

                             do  until rsx.eof

                               xvale = "N"


                               StrSql = "Select * From REQUERIMIENTO where codigo = '" & RSX(4) & "' "
                               set rsxp = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                               if NOT rsxp.eof then
                                  xCLI = rsxp(2) 
                                  xCLIN = rsxp(0) 
                                  xvale = "S"
                               end if

                               call ofv.cerrarconsulta(rsxp)    
    
                               IF  xvale = "S" THEN 

                                 xnohay2 = "N"

                                 INTabla = INTabla + 1


                                 xarc = "<font face=webdings color=#993333 size=4  >&#0157;</font>"

                                 sHref     = "&nbsp;"
                                 sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('editarPRY.asp?codigo=" &  RSX(4) & "&tipo=REQ&SEGCKTPSC=" & RSX(15) & "'); return false;" & Chr(34)

                                 wNomL = "&nbsp;REQUERIMIENTO: <b>"   & xCLIN & "-" & XCLI & "</b>&nbsp;&nbsp;&nbsp;CIRCUITO: <b>" & RSX(7) & "</b>&nbsp;&nbsp;&nbsp;PASO: <b>"  & RSX(31) & "</b>"

                                 sImagen   = wNomL

                                 if xarc <> "NO" then
                                    sImagen = xarc & sImagen
                                 end if

                                 sLink     = OFV.GenLink("aaa title='Autorizar circuito ...' ", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")

                                 xfromusr = rsx(6)


                                 xxxx = getInfoUser(rsx(6))
                                 xfromusr = session("SGU_nombre")



                                 ANTabla(INTabla,0) = "CKT" & INTabla
                                 ANTabla(INTabla,1) = 5
                                 ANTabla(INTabla,2) = 0
                                 ANTabla(INTabla,3) = rsx(33)
                                 ANTabla(INTabla,4) = rsx(34)
                                 ANTabla(INTabla,5) = "S"
                                 ANTabla(INTabla,6) = "N"
                                 ANTabla(INTabla,7) = xfromusr
                                 ANTabla(INTabla,8) = sLink

 
                              end if


                             rsx.movenext
                             loop
                        end if
                        call ofv.cerrarconsulta(rsx)




       case 3

                        session("ModMain") = "CL"




                         SEGWHERE = " WHERE OBJTIPO = '" & session("OBJTIPO") & "' AND  FIN = 'NO' AND  ( (RESPASG = '" & Session("Usuario") & "') OR ("

                         xxxx     = getInfoUser(Session("Usuario"))
                         xusrlog  = session("SGU_nombre")


                         if  xusrlog <> "NO" THEN


                             SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "') ) "
                             SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "


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





                             if  cstr(session("SGU_idcostcenter") & "") <> "0" then
                                 SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
                                 SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & session("SGU_idcostcenter") & "' OR SUBCODIGO = '" & session("SGU_idcostcenter") & "' ) "
                               ''''  SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
                                 SEGWHERE = SEGWHERE & " ) ) "
                             END IF


                         END IF

                         SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "

                         Strsql = "SELECT * FROM CKTVIEW " & SEGWHERE

                      ''' RESPONSE.WRITE STRSQL & "<BR>"
                         set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         if  rsx.eof then
                             xnohay2 = "S"
                         else

                             do  until rsx.eof

                               xvale = "N"


                               StrSql = "Select * From REQUERIMIENTO where codigo = '" & RSX(4) & "' "
                               set rsxp = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                               if NOT rsxp.eof then
                                  xCLI = rsxp(2) 
                                  xCLIN = rsxp(0) 
                                  xvale = "S"
                               end if

                               call ofv.cerrarconsulta(rsxp)    
    
                               IF  xvale = "S" THEN 

                                 xnohay2 = "N"

                                 INTabla = INTabla + 1


                                 xarc = "<font face=webdings color=#005500 size=4  >&#0157;</font>"

                                 sHref     = "&nbsp;"
                                 sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('editarPRY.asp?codigo=" &  RSX(4) & "&tipo=REQ'); return false;" & Chr(34)

                                 wNomL = "&nbsp;REQUERIMIENTO: <b>"  & xCLIN & "-"  & XCLI & "</b>&nbsp;&nbsp;&nbsp;CIRCUITO: <b>" & RSX(7) & "</b>&nbsp;&nbsp;&nbsp;PASO: <b>"  & RSX(16) & "</b>"

                                 sImagen   = wNomL

                                 if xarc <> "NO" then
                                    sImagen = xarc & sImagen
                                 end if

                                 sLink     = OFV.GenLink("aaa title='Consultar circuito ...' ", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")

                                 xfromusr = rsx(6)


                                 xxxx = getInfoUser(rsx(6))
                                 xfromusr = session("SGU_nombre")



                                 ANTabla(INTabla,0) = "CKT" & INTabla
                                 ANTabla(INTabla,1) = 2
                                 ANTabla(INTabla,2) = 0
                                 ANTabla(INTabla,3) = rsx(10)
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


'''''''''''''''''''''''''''''''''''''''''''''


       case 4

                     session("ModMain") = "CL"
            
                        


                         fdesde = clng(date()) - 8
                         fhasta = clng(date()) + 1

                         Strsql = "SELECT distinct codigo FROM PROYtskasgn where usrresp = '" & Session("Usuario") & "' and estado = 1 and vplan = pvplan  "


                       ''  response.write strsql
                         set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)





                         if  rsx.eof then
                             xnohay2 = "S"
                         else
                         XPRYANT = "NO"
                         do until rsx.eof

                            IF XPRYANT <> RSX(0) THEN

                               XPRYANT = RSX(0)


                                 xnohay2 = "N"

                                 INTabla = INTabla + 1



                                StrSql = "Select * From proyecto where codigo = '" & RSX(0) & "' "
                                set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                if NOT rs.eof then
                                   xCLI = RS(2) 
                                   xfromusr = rs(14)
                                   xdescri = rs(10)
                                   xfechasol = rs(4)
                                end if



                          
                                wNomL = "&nbsp;Proyecto: <b>" & XCLI & "</b>&nbsp;<br>&nbsp;<b>" & xdescri & "</b>"

                                 xarc = "<font face=webdings color=#6699CC size=4  >&#0157;</font>"

                                 sHref     = "&nbsp;"
                                 sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('editarPRY.asp?codigo=" &  RSX(0) & "&tipo=PRY'); return false;" & Chr(34)

                                 sImagen   = wNomL

                                 if xarc <> "NO" then
                                    sImagen = xarc & sImagen
                                 end if

                                 sLink     = OFV.GenLink("aaa title='Consultar proyecto ...' ", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")

 
                                 xxxx = getInfoUser(xfromusr)
                                 xfromusr = session("SGU_nombre")



                                 ANTabla(INTabla,0) = "CKT" & INTabla
                                 ANTabla(INTabla,1) = 4
                                 ANTabla(INTabla,2) = 0
                                 ANTabla(INTabla,3) = dFechai
                                 ANTabla(INTabla,4) = time()
                                 ANTabla(INTabla,5) = "S"
                                 ANTabla(INTabla,6) = "N"
                                 ANTabla(INTabla,7) = xfromusr
                                 ANTabla(INTabla,8) = sLink

 
                              end if


                             rsx.movenext
                             loop
                        end if
                        call ofv.cerrarconsulta(rsx)



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        '''' ESTADO DE TAREAS EN CURSO DEL PROYECTO



                         Strsql = "SELECT * FROM PROYTSKSTATUS where ASIGNADO = '" & Session("Usuario") & "' and ESTADO = 1  AND FECHAEREAL = 0  order by fechaEPLAN"


                         ''response.write strsql
                         set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)




                         if  rsx.eof then
                             xnohay2 = "S"
                         else

                         do until rsx.eof


                                 xnohay2 = "N"

                                 INTabla = INTabla + 1


                            StrSql = "Select * From proyecto where codigo = '" & RSX(0) & "' "
                            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                            if NOT rs.eof then
                               xCLI = RS(2) 
                               xUSR = RS(14) 
                            end if



                          
                                wNomL = "&nbsp;Proyecto: <b>" & XCLI & "</b>&nbsp;<br>&nbsp;<b><u> TAREA: " & RSX(3) & "-" & RSX(4) & "</u> PRESENTA UN DESVIO RESPECTO DEL PLAN DEL PROYECTO.</b>"


                                 xarc = "<font face=webdings color=#F7AA22 size=4  >&#0157;</font>"

                                 sHref     = "&nbsp;"
                                 sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('editarPRY.asp?codigo=" &  RSX(0) & "&tipo=PRY'); return false;" & Chr(34)

                                 sImagen   = wNomL

                                 if xarc <> "NO" then
                                    sImagen = xarc & sImagen
                                 end if

                                 sLink     = OFV.GenLink("aaa title='Consultar proyecto ...' ", "cmb", sHref, "", "", sFuncion, "<font color=#999999 >" & sImagen & "</font>", "si")

                                 xfromusr = xUSR
                                 xxxx = getInfoUser(xUSR)
                                 xfromusr = session("SGU_nombre")



                                 ANTabla(INTabla,0) = "CKT" & INTabla
                                 ANTabla(INTabla,1) = 3
                                 ANTabla(INTabla,2) = 0
                                 ANTabla(INTabla,3) = dFechai
                                 ANTabla(INTabla,4) = time()
                                 ANTabla(INTabla,5) = "S"
                                 ANTabla(INTabla,6) = "N"
                                 ANTabla(INTabla,7) = xfromusr
                                 ANTabla(INTabla,8) = sLink

 



                             rsx.movenext
                             loop
                        end if
                        call ofv.cerrarconsulta(rsx)



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 

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

 IF  Session("FULLGUI") = "N" THEN
     shtml = shtml & "background-color:#eeeeee;"
 else     
     shtml = shtml & XspEx("#eeeeee",Session("CMFPNTIT"),0) & ";"
 end if

 shtml = shtml & "' >"



     sHTML   = sHTML &  "<tr >"

     sHTML   = sHTML &  "<td class='UT'  >"

             sHTML = sHTML &  "<table  cellspacing=1 cellpadding=3  >"

             for ii=2 to wmntope
                 if  atit(ii) <> "NO" then

                    if  ii = 10 then
                   '''' if  ii = menu or ii = 10 then
                        stilo  = stilo2
                    else

                        if  ii = menu  then
                            xsclase = "nt STYLE='border-style:inset;border-color:" & Session("CMFPNTITFT") & ";border-width:1;padding-left:5px;padding-right:5px;" & XspEx("#f7aa22","#dddddd",0) & ";' "
                        else
                            xsclase = "nt STYLE='border-right-style:solid;border-right-color:" & Session("CMFPNTITFT") & ";border-right-width:1;padding-left:5px;padding-right:5px;' "
                        end if 

                        xmDato = REPLACE(atit(ii)," ","&nbsp;")

                        if ii < 11 then
                           xmhref = "<a class='xts' href='SEGNOVEDCIRCUITO.asp?menu=" & ii & "' >" & xmDato & "</a>"

                        else
                           xmhref = "<a class='xts' href='" & awmnasp(ii) & "' >" & xmDato & "</a>"


                        end if

                        sHTML  = sHTML &  "<td class=" & xsclase & " align=center  >"
                        sHTML  = sHTML &  xmhref & "</TD>"
                    end if

                 end if
            next

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
                      xtipomes = "CIRCUITOS"
                 case 3
                      xtipomes = "TAREAS"
                 case 4
                      xtipomes = "PROYECTOS"
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

        '  if  xtfvig = dFechai then
        '      XMARQ  = sglmar
        '  else
        '      XMARQ  = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
        '  end if

          IF  xtipom = 5 THEN
              xtfechavig = fechaArg(cdate(xtfvig + clng(cdate("01-01-1900"))))  & "&nbsp;-&nbsp;" & GETGLBHMSW(xthvig) 'xthvig
          ELSE
              xtfechavig = xtfvig 'fechaArg(cdate(xtfvig + clng(cdate("01-01-1900")))) & "&nbsp;&nbsp;" & xthvig
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
                         sHTML = sHTML &  "<td class=AAA colspan=8 style='line-height:2pt;' >&nbsp;</TD>"
                         sHTML = sHTML &  "</TR>"

                     if xtmcod <> "0" then

                         sHTML = sHTML &  "<tr bgcolor=" & xTrColor & " >"
                         sHTML = sHTML &  "<td class=" & xfclas & " width=1% bgcolor=#FFFFFF >" & sImage & "</TD>"
                         sHTML = sHTML &  "<td class=vt width=55%  STYLE='text-align:justify;background-color:#DDDDDD;filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & Session("CMFPNTIT") & ", endColorStr=" & xTrColor & ", gradientType=1);' >&nbsp;&nbsp;<a class=tt >" & xtfechavig & "</a></TD>"
                         sHTML = sHTML &  "<td class=ut width=13% >&nbsp;</TD>"
                         sHTML = sHTML &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtipomes & "</font></TD>"
                         sHTML = sHTML &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtmcod & "</font></TD>"
                         sHTML = sHTML &  "<td class=btba width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnadj & "</TD>"
                         sHTML = sHTML &  "<td class=btbn width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnlnk & "</TD>"
                         sHTML = sHTML &  "<td class=qt width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & XMARQ & "</TD>"
                         sHTML = sHTML &  "</TR>"

                     end if

                         sHTML = sHTML &  "<tr bgcolor=" & xTrColor & " >"
                         sHTML = sHTML &  "<td class=uti colspan=8 STYLE='text-align:justify;' ><br><font color=#999999 >" & xtmsg & "</font></TD>"
                         sHTML = sHTML &  "</TR>"


                     if xtmcod <> "0" then


                         sHTML = sHTML &  "<tr bgcolor=" & xTrColor & " height=20 >"
                         sHTML = sHTML &  "<td class=ut colspan=8 valign=top align=right style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' >" & sArma & "</TD>"
                         sHTML = sHTML &  "</TR>"

                     end if


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF

      next


  end if



  sHTML = sHTML &  "</table>"


 shtml = shtml & "</div>"

  sHTML = sHTML &  "</center>"

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

                 if  Session("DELLOG") = "S" then

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


%>
