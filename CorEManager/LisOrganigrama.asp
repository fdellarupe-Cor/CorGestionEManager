<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

sub linea(xfila,xindi)
    if xindi > 40 then
       xindi = 0
       response.write "</table>"
       response.write "</td>"
       response.write "</tr>"
       response.write "</table>"

       if ofv.explorador = "MSIE" then response.write "<div align=center > &nbsp </div>"

       response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>" & chr(13)
       response.write "<tr>"
       response.write "<td width=100% >"
       response.write ofv.RepHeader2(ucase("Lista del Organigrama"),"600")

       response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
       response.write "<tr>"
       response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
       response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
       response.write "</tr>"
    end if

response.write xfila
xindi = xindi + 1
end sub

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0

    indi = 0
       response.write "<html> <link rel=""stylesheet"" type=""text/css"" href=""../fonts/Corfont/Corflaticon.css"" >" & corstyler() & "<body>"
       response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
       response.write "<tr>"
       response.write "<td width=100% >"
       response.write  ofv.RepHeader2(ucase("Lista del Organigrama"),"600")
       response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
       response.write "<tr>"
       response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
       response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
       response.write "</tr>"

 scolor = "<font face=corflaticon size=2 color=#336699>4</font>&nbsp;"
 scolor2 = "<font face=corflaticon size=2 color=#993333>a</font>&nbsp;"

dim a1(200)
dim a2(200)
dim a3(200)
dim a4(200)
dim b1(200)
dim b2(200)
dim b3(200)
dim b4(200)

for iix=0 to 200
    a1(iix) = "xxx"
    a2(iix) = "xxx"
    a3(iix) = "xxx"
    a4(iix) = "xxx"
    b1(iix) = "xxx"
    b2(iix) = "xxx"
    b3(iix) = "xxx"
    b4(iix) = "xxx"
next

iix=0

 set cn = ofv.conectar(ofv.strconn2)

'response.write "STRING: " & request.querystring & "<br>" & "FORM: " & request.form


    Strsql = "select * from csOrgmaster where ecodepend is null and ecoun = " & SESSION("UN")

    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    do until rs.eof
       b1(iix) = 0
       b2(iix) = "SD"
       b3(iix) = rs(0)
       b4(iix) = rs(1)
    ''   response.write rs(1) & "<br>"
       iix = iix + 1
    rs.movenext
    loop
    b1(iix) = "xxx"

xsale = "SSS"
do while xsale = "SSS"


         for iix = 0 to 200
             if b1(iix) = "xxx" then
                if iix = 0 then xsale = "xxx"
                a1(iix) = "xxx"
                exit for
             end if
             a1(iix) = b1(iix)
             a2(iix) = b2(iix)
             a3(iix) = b3(iix)
             a4(iix) = b4(iix)
         next
         b1(0) = "xxx"
         iib = 0

         for iix = 0 to 200
             if a1(iix) = "xxx" then
                exit for
             end if
         strsql = "select ecorel.ecomaster,ecorel.ecodepend,eco.descri from "
         strsql = strsql & " ecorel, eco where ecorel.ecodepend = eco.eco "
         strsql = strsql & " and ecorel.ecomaster = " & a3(iix) & " and eco.ecoun = " & SESSION("UN")
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

        ' if not rs.eof then
            subtemax = "<tr><td>&nbsp;</td><td>&nbsp;</td>"
            subtemax = subtemax & "<td>&nbsp;</td><td>&nbsp;</td></tr>"
            call linea(subtemax,indi)

            subtemax = "<tr><td class=cc>" & scolor2 & "</td><td class=cc colspan=2>"
            subtemax = subtemax & "<b>" & a4(iix) & "</b>"
            subtemax = subtemax & "</td><td class=cc >"
            if a2(iix) = "SD" then
               subtemax = subtemax & "&nbsp;"
            else
               subtemax = subtemax & "<b>DEPENDENCIA: " & a2(iix)
            end if
            subtemax = subtemax & "</b></td></tr>"
            call linea(subtemax,indi)
       '  end if

         do until rs.eof

            b1(iib) = a3(iix)
            b2(iib) = a4(iix)
            b3(iib) = rs(1)
            b4(iib) = rs(2)
            iib = iib + 1

            subtemax = "<tr><td>&nbsp;</td><td>&nbsp;</td><td class=ut colspan=2 >"
            subtemax = subtemax & scolor & rs(2) & "</td>"
            subtemax = subtemax & "</tr>"
            call linea(subtemax,indi)

         rs.movenext
         loop
         next
         b1(iib) = "xxx"
    loop
    call ofv.cerrarconn(cn)

  response.write "</table>"
  response.write "</td>"
  response.write "</tr>"
  response.write "</table>"
  response.write "<br><br></body></html>"
end if

response.write sHTML

%>
