<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

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

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


stilo1 = " style={background-color:" & session("CMFMENU") & ";border-style:outset;border-width:2;}"
stilo2 = " style={background-color:" & session("CMFMNBOT") & ";}"

 OBJTIPO = request.querystring("OBJTIPO")
 if OBJTIPO = "" then
    OBJTIPO = session("OBJTIPO")
 else
    session("OBJTIPO") = OBJTIPO
 end if 

set cn = ofv.conectar(ofv.strconn2)
set cn0 = ofv.conectar(ofv.strconn0)
set cn1 = ofv.conectar(ofv.strconn5)
set cn4 = ofv.conectar(ofv.strconn6)

dim atit(10)
titc = 0
Strsql = "SELECT * FROM novedtit "
set rsx = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
if not rsx.eof then

   for tt = 2 to 10
       atit(tt) = rsx(tt)
       if rsx(tt) <> "NO" then titc = titc + 1
   next


end if
call ofv.cerrarconsulta(rsx)




menu = request.querystring("menu")
if menu = "" then menu = 2
menu = cint(menu)





  sHTML = replace(ofv.menuheader("","novedades "),"<html>","<html><META http-equiv='refresh' content='10' >")

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"



    sHTML = sHTML & xs & s

'  sHTML =   ofv.FormBody

  sHTML = sHTML &  "<table bgcolor=#ffffff width=100% cellspacing=0 cellpadding=2 style={border-style:outset;border-width:2;} >"

  XMARQ = "<MARQUEE DIRECTION=left BEHAVIOR=scroll SCROLLAMOUNT=5 SCROLLDELAY=200  >"
  XMARQ = XMARQ & "<font face=wingdings color=#ffcc44 size=2>&#0184;&#0185;&#0186; </font>&nbsp;NOVEDADES: " & atit(menu) & " AL "  & now() & "&nbsp;<font face=wingdings color=#ffc544 size=2> &#0184;&#0185;&#0186;</font></MARQUEE>"



  sHTML = sHTML &  "<tr>"
  sHTML = sHTML &  "<td class=vt width=40% style={border-style:inset;border-width:2;} ><b>"  & XMARQ & "</b></TD>"
for ii=2 to 10
    if atit(ii) <> "NO" then
         if ii = menu then
            stilo = stilo2
         else
            stilo = stilo1
         end if
         sHTML = sHTML &  "<td " & stilo & " align=center ><a href=SEGNOVEDCIRCUITO.asp?menu=" & ii & "><font face=tahoma size=1 color=#ffffff>&nbsp;<b>" & atit(ii) & "</b></font></a></TD>"
    end if
next

  sHTML = sHTML &  "</TR>"

  sHTML = sHTML &  "<tr>"
  sHTML = sHTML &  "<td " & stilo2 & " colspan=" & (titc + 1) & "  ><hr size=1 color=#ffffff></TD>"


  sHTML = sHTML &  "</table>"


  sHTML = sHTML &  "<table bgcolor=#ffffff width=100% cellspacing=0 cellpadding=2  style={border-style:outset;border-width:2;}  >"



nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

        dFechai = clng(date()) - nFecini

select case menu

       case 1

       case 2
  session("ModMain") = "WK"
                        
                         SEGWHERE = " WHERE OBJTIPO = '" & session("OBJTIPO") & "' AND  PSCFIN = 'NO' AND ( (RESPASG = '" & Session("Usuario") & "') OR ("

                         StrSql = "Select * From USUARIOS WHERE ID = '" & Session("Usuario") & "' "
                         set rs = ofv.crearconsultaEx(StrSql,cn0,1,parametros)
                         if not rs.eof THEN

                            SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "') ) "  
                            SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND (PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "  
                            IF RS(21) > 0 THEN
                               SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
                               SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & RS(21) & "' OR SUBCODIGO = '" & RS(21) & "' ) "
                               SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
                               SEGWHERE = SEGWHERE & " ) ) "  
                            END IF


                         END IF                          
                         SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "  

                         Strsql = "SELECT * FROM CKTVIEW " & SEGWHERE

                       ''RESPONSE.WRITE STRSQL & "<BR>" 
                         set rsx = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
                         if rsx.eof then

                     sHTML = sHTML &  "<tr>"
                     sHTML = sHTML &  "<td class=VT width=40% ><b>no hay &nbsp;novedades de " & atit(menu) & " <br></b></TD>"
                     sHTML = sHTML &  "</TR>"
                         else
                         do until rsx.eof



                            xlink = "<A class=cc href='&nbsp;'  "
                            xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" &  RSX(4) & "&tipo=REQ&SEGCKTPSC=" & RSX(8) & "'); return false;" & chr(34)
                            xlink = xlink & " ><IMG  src='Imagenes/ICorAbierto.jpg'  border=0 ></A>"


                            StrSql = "Select * From REQUERIMIENTO where codigo = '" & RSX(4) & "' "
                            set rs = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
                            if NOT rs.eof then
                               xCLI = RS(2) 
                            end if



                          
                           wNomL = "&nbsp;REQUERIMIENTO: <b>" & XCLI & "</b>&nbsp;&nbsp;&nbsp;CIRCUITO: <b>" & RSX(7) & "</b>&nbsp;&nbsp;&nbsp;PASO: <b>"  & RSX(16) & "</b>"

                            xlink = xlink & "<A class=cc href='&nbsp;' onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" &  RSX(4) & "&tipo=REQ&SEGCKTPSC=" & RSX(8) & "'); return false;" & chr(34) & " >" &  wNomL & "</A>"


 
                           sLink = xlink


                               sxopc4 = "<font color=#800000 >ALTA</font>"
                               sxopc4x = "<font color=#800000 size=4><b>!</b></font>"


                               xnove = "&nbsp;"
                               XMARQ = "&nbsp;"


                         xfromusr = rsx(6)
                         strsql = "Select descripcion From Usuarios  where ID = '" & rsx(6) & "' "
                         set rsx1 = ofv.crearconsultaEx(StrSql,cn0,1,volver)
                         if not rsx1.eof then
                               xfromusr = rsx1(0)
                         end if
                         call ofv.cerrarconsulta(rsx1)


                            sHTML = sHTML &  "<tr bgcolor=#ddeeff>"
                            sHTML = sHTML &  "<td class=cc  width=1% ><b>" & sxopc4x & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=20% >Vigencia: <b>" & rsx(10) & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=20% >Importancia: <b>" & sxopc4 & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=27% >Codigo: <b>" & rsx(8) & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=8% >" & XMARQ & "</TD>"
                            sHTML = sHTML &  "<td class=ut  width=5% >" & xnove & "</TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr  >"
                            sHTML = sHTML &  "<td class=ut colspan=6  ><br><font face=wingdings color=#800000 size=5><b>P</b></font> " & sLink & "</TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr  >"
                            sHTML = sHTML &  "<td class=ut colspan=6 align=right ><u>Origen:&nbsp;" & xfromusr & "</u></TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr >"
                            sHTML = sHTML &  "<td class=ut colspan=6  ><hr size=1 color=#99ccff></b></TD>"
                            sHTML = sHTML &  "</TR>"




                         rsx.movenext
                         loop
                        end if
                        call ofv.cerrarconsulta(rsx)

       case 3

                          session("ModMain") = "CL"
                        
                         SEGWHERE = " WHERE OBJTIPO = '" & session("OBJTIPO") & "' AND  FIN = 'NO' AND  ( (RESPASG = '" & Session("Usuario") & "') OR ("

                         StrSql = "Select * From USUARIOS WHERE ID = '" & Session("Usuario") & "' "
                         set rs = ofv.crearconsultaEx(StrSql,cn0,1,parametros)
                         if not rs.eof THEN
                            SEGWHERE = SEGWHERE & " ((RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "')) "  
                            SEGWHERE = SEGWHERE & " OR ((RESPTIPO = 'GRP') AND (PSCRESPONSABLE = '" & SESSION("SEGGRP") & "')) "  
                            IF RS(21) > 0 THEN
                            SEGWHERE = SEGWHERE & " OR ((RESPTIPO = 'SEC') AND (PSCRESPONSABLE = '" & RS(21) & "')) "  
                            END IF


                         END IF                          
                         SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "  

                         Strsql = "SELECT * FROM CKTVIEW " & SEGWHERE

                         set rsx = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
                         if rsx.eof then

                     sHTML = sHTML &  "<tr>"
                     sHTML = sHTML &  "<td class=VT width=40% ><b>no hay &nbsp;novedades de " & atit(menu) & " <br></b></TD>"
                     sHTML = sHTML &  "</TR>"
                         else
                         do until rsx.eof



                            xlink = "<A class=cc href='&nbsp;'  "
                            xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" &  RSX(4) & "&tipo=REQ'); return false;" & chr(34)
                            xlink = xlink & " ><IMG  src='Imagenes/ICorAbierto.jpg'  border=0 ></A>"


                            StrSql = "Select * From REQUERIMIENTO where codigo = '" & RSX(4) & "' "
                            set rs = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
                            if NOT rs.eof then
                               xCLI = RS(2) 
                            end if



                          
                           wNomL = "&nbsp;REQUERIMIENTO: <b>" & XCLI & "</b>&nbsp;&nbsp;&nbsp;CIRCUITO: <b>" & RSX(7) & "</b>&nbsp;&nbsp;&nbsp;PASO: <b>" & RSX(16) & "</b>"

                            xlink = xlink & "<A class=cc href='&nbsp;' onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" &  RSX(4) & "&tipo=REQ'); return false;" & chr(34) & " >" &  wNomL & "</A>"


 
                           sLink = xlink


                               sxopc4 = "<font color=#800000 >ALTA</font>"
                               sxopc4x = "<font color=#800000 size=4><b>!</b></font>"


                               xnove = "&nbsp;"
                               XMARQ = "&nbsp;"


                         xfromusr = rsx(6)
                         strsql = "Select descripcion From Usuarios  where ID = '" & rsx(6) & "' "
                         set rsx1 = ofv.crearconsultaEx(StrSql,cn0,1,volver)
                         if not rsx1.eof then
                               xfromusr = rsx1(0)
                         end if
                         call ofv.cerrarconsulta(rsx1)


                            sHTML = sHTML &  "<tr bgcolor=#ddeeff>"
                            sHTML = sHTML &  "<td class=cc  width=1% ><b>" & sxopc4x & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=20% >Vigencia: <b>" & rsx(10) & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=20% >Importancia: <b>" & sxopc4 & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=27% >Codigo: <b>" & rsx(8) & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=8% >" & XMARQ & "</TD>"
                            sHTML = sHTML &  "<td class=ut  width=5% >" & xnove & "</TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr  >"
                            sHTML = sHTML &  "<td class=ut colspan=6  ><br><font face=wingdings color=#800000 size=5><b>P</b></font> " & sLink & "</TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr  >"
                            sHTML = sHTML &  "<td class=ut colspan=6 align=right ><u>Origen:&nbsp;" & xfromusr & "</u></TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr >"
                            sHTML = sHTML &  "<td class=ut colspan=6  ><hr size=1 color=#99ccff></b></TD>"
                            sHTML = sHTML &  "</TR>"




                         rsx.movenext
                         loop
                        end if
                        call ofv.cerrarconsulta(rsx)


       case 4

  session("ModMain") = "CL"
                     sHTML = sHTML &  "<tr>"
                     sHTML = sHTML &  "<td class=cc   width=40% colspan=6 ><b>" & atit(menu) & " CON TAREAS ASIGNADAS A:&nbsp;" & SESSION("SEGUSUARIO") & "<br></b></TD>"
                     sHTML = sHTML &  "</TR>"                        
                        


                         fdesde = clng(date()) - 8
                         fhasta = clng(date()) + 1

                    '     Strsql = "SELECT * FROM PROYNOVE where usrafc = '" & Session("Usuario") & "' and fechasol > " & fdesde & " and fechasol < " & fhasta & "   order by fechasol, CODIGO"

                        Strsql = "SELECT distinct codigo FROM PROYtskasgn where usrresp = '" & Session("Usuario") & "' and estado = 1 and vplan = pvplan  "


                       ''  response.write strsql
                         set rsx = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
                         if rsx.eof then

                     sHTML = sHTML &  "<tr>"
                     sHTML = sHTML &  "<td class=VT  bgcolor=#f7aa22  width=40% colspan=6 ><b>no hay &nbsp;novedades de " & atit(menu) & " iniciados<br></b></TD>"
                     sHTML = sHTML &  "</TR>"
                         else
                         XPRYANT = "NO"
                         do until rsx.eof

                            IF XPRYANT <> RSX(0) THEN

                            XPRYANT = RSX(0)

                            xlink = "<A class=cc href='&nbsp;'  "
                            xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" &  RSX(0) & "&tipo=PRY'); return false;" & chr(34)
                            xlink = xlink & " ><IMG  src='Imagenes/ICorAbierto.jpg'  border=0 ></A>"


                            StrSql = "Select * From proyecto where codigo = '" & RSX(0) & "' "
                            set rs = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
                            if NOT rs.eof then
                               xCLI = RS(2) 
                               xfromusr = rs(14)
                               xdescri = rs(10)
                               xfechasol = rs(4)
                            end if



                          
                           wNomL = "&nbsp;Proyecto: <b>" & XCLI & "</b>&nbsp;<br>&nbsp;<b>" & xdescri & "</b>"

                            xlink = xlink & "<A class=cc href='&nbsp;' onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" &  RSX(0) & "&tipo=PRY'); return false;" & chr(34) & " >" &  wNomL & "</A>"


 
                           sLink = xlink


                               sxopc4 = "<font color=#800000 >ALTA</font>"
                               sxopc4x = "<font color=#800000 size=4><b>!</b></font>"


                               xnove = "&nbsp;"
                               XMARQ = "&nbsp;"



                         strsql = "Select descripcion From Usuarios  where ID = '" & xfromusr & "' "
                         set rsx1 = ofv.crearconsultaEx(StrSql,cn0,1,volver)
                         if not rsx1.eof then
                               xfromusr = rsx1(0)
                         end if
                         call ofv.cerrarconsulta(rsx1)


                            sHTML = sHTML &  "<tr bgcolor=#ddeeff>"
                            sHTML = sHTML &  "<td class=cc  width=1% ><b>" & sxopc4x & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=20% >Vigencia: <b>" & cdate(xfechasol + clng(cdate("01-01-1900"))) & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=20% >Importancia: <b>" & sxopc4 & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=27% >Codigo: <b>" & rsx(0) & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=8% >" & XMARQ & "</TD>"
                            sHTML = sHTML &  "<td class=ut  width=5% >" & xnove & "</TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr  >"
                            sHTML = sHTML &  "<td class=ut colspan=6  ><br><font face=wingdings color=#800000 size=5><b>P</b></font> " & sLink & "</TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr  >"
                            sHTML = sHTML &  "<td class=ut colspan=6 align=right ><u>Origen:&nbsp;" & xfromusr & "</u></TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr >"
                            sHTML = sHTML &  "<td class=ut colspan=6  ><hr size=1 color=#99ccff></b></TD>"
                            sHTML = sHTML &  "</TR>"

                           END IF


                         rsx.movenext
                         loop
                        end if
                        call ofv.cerrarconsulta(rsx)

                     sHTML = sHTML &  "<tr>"
                     sHTML = sHTML &  "<td     colspan=6 >&nbsp;</TD>"
                     sHTML = sHTML &  "</TR>"   





                        '''' ESTADO DE TAREAS EN CURSO DEL PROYECTO

                     sHTML = sHTML &  "<tr>"
                     sHTML = sHTML &  "<td class=cc   width=40% colspan=6 ><b>ESTADO DE TAREAS EN CURSO DE " & atit(menu) & "<br></b></TD>"
                     sHTML = sHTML &  "</TR>"   


                         Strsql = "SELECT * FROM PROYTSKSTATUS where ASIGNADO = '" & Session("Usuario") & "' and ESTADO = 1  AND FECHAEREAL = 0  order by fechaEPLAN"


                         ''response.write strsql
                         set rsx = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
                         if rsx.eof then

             '        sHTML = sHTML &  "<tr>"
             '        sHTML = sHTML &  "<td class=VT width=40% ><b>no hay &nbsp;novedades de TAREAS DE " & atit(menu) & " <br></b></TD>"
             '        sHTML = sHTML &  "</TR>"
                         else
                         do until rsx.eof



                            xlink = "<A class=cc href='&nbsp;'  "
                            xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" &  RSX(0) & "&tipo=PRY'); return false;" & chr(34)
                            xlink = xlink & " ><IMG  src='Imagenes/ICorAbierto.jpg'  border=0 ></A>"


                            StrSql = "Select * From proyecto where codigo = '" & RSX(0) & "' "
                            set rs = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
                            if NOT rs.eof then
                               xCLI = RS(2) 
                               xUSR = RS(14) 
                            end if



                          
                           wNomL = "&nbsp;Proyecto: <b>" & XCLI & "</b>&nbsp;<br>&nbsp;<b><u> TAREA: " & RSX(3) & "-" & RSX(4) & "</u> PRESENTA UN DESVIO RESPECTO DEL PLAN DEL PROYECTO.</b>"

                            xlink = xlink & "<A class=cc href='&nbsp;' onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" &  RSX(0) & "&tipo=PRY'); return false;" & chr(34) & " >" &  wNomL & "</A>"


 
                           sLink = xlink


                               sxopc4 = "<font color=#800000 >ALTA</font>"
                               sxopc4x = "<font color=#800000 size=4><b>!</b></font>"


                               xnove = "&nbsp;"
                               XMARQ = "&nbsp;"


                         xfromusr = xUSR
                         strsql = "Select descripcion From Usuarios  where ID = '" & xUSR & "' "
                         set rsx1 = ofv.crearconsultaEx(StrSql,cn0,1,volver)
                         if not rsx1.eof then
                               xfromusr = rsx1(0)
                         end if
                         call ofv.cerrarconsulta(rsx1)

                         xfromusr = "Administrador de Proyectos - Cor OnTime"


                            sHTML = sHTML &  "<tr bgcolor=#ddeeff>"
                            sHTML = sHTML &  "<td class=cc  width=1% ><b>" & sxopc4x & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=20% >Vigencia: <b>" & cdate(rsx(9) + clng(cdate("01-01-1900"))) & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=20% >Importancia: <b>" & sxopc4 & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=27% >Codigo: <b>" & rsx(0) & "</b></TD>"
                            sHTML = sHTML &  "<td class=ut  width=8% >" & XMARQ & "</TD>"
                            sHTML = sHTML &  "<td class=ut  width=5% >" & xnove & "</TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr  >"
                            sHTML = sHTML &  "<td class=ut colspan=6  ><br><font face=wingdings color=#800000 size=5><b>P</b></font> " & sLink & "</TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr  >"
                            sHTML = sHTML &  "<td class=ut colspan=6 align=right ><u>Origen:&nbsp;" & xfromusr & "</u></TD>"
                            sHTML = sHTML &  "</TR>"

                            sHTML = sHTML &  "<tr >"
                            sHTML = sHTML &  "<td class=ut colspan=6  ><hr size=1 color=#99ccff></b></TD>"
                            sHTML = sHTML &  "</TR>"




                         rsx.movenext
                         loop
                        end if
                        call ofv.cerrarconsulta(rsx)




     end select

     call ofv.cerrarconn(cn)
     call ofv.cerrarconn(cn0)
     call ofv.cerrarconn(cn1)

  sHTML = sHTML &  "</table>"
  sHTML = sHTML &  "</body></html>"



  response.write sHTML
end if

%>
