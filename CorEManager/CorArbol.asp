<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


Function GeneraResto(vecsec(), vecsecd(),xlevel)

fila = ""



For xin = 1 To UBound(vecsec)

    bordert = ""
    borderl = ""
    borderr = ""
    bstilo = ""
    bcolor = ""



    xcolspan  = (2 * (xlevel - 2)) + 1
    xcolspan2 = 22 - xcolspan


           celda = OFV.gencelda("", "aaa colspan=" & xcolspan & " ", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa colspan=" & xcolspan2 & " style='" & cborderl & "'", "", "", "", "", "", "&nbsp;", "si")
           fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")


           celda = OFV.gencelda("", "aaa colspan=" & xcolspan & " rowspan=2 ", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa  style='" & cborderl & cborderb & "'", "", "", "", "", "", "&nbsp;", "si")

           celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & cstilo & "'", "center", "", "", "", "",simagenE0, "si")




                shref = "corarbol.asp?ID=" & vecsec(xin) & "&Sector=" & OFV.ConvertirCar(vecsecd(xin), " ", "_")
                sinput = "<font face=tahoma color=#000080 size=1><b>" & vecsecd(xin) & "</b></font>"
                slink = OFV.GenLink("", "cc", shref, "", "", "", sinput, "si")


           celda = celda & OFV.gencelda("", "aaa colspan=" & (21 - (xcolspan + 3)) & " rowspan=2 ", "", "", "", "", "","&nbsp;" & slink, "si")
           celda = celda & OFV.gencelda("", "aaa ", "", "", "", "", "", "&nbsp;", "si")
           fila  = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

      if xin = UBound(vecsec) then
         borderl = ""
      else
         borderl = cborderl
      end if

           celda = OFV.gencelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa ", "", "", "", "", "", "&nbsp;", "si")
           fila  = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")





Next


GeneraResto = fila


End Function


ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
      Generar = ""

    sImagenE0 = "<B><FONT FACE=corflaticon COLOR=#6699CC SIZE=3>4</FONT></B>"
    sImagen4si = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>&#0128;</FONT></B>"
    sImagenin = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>s</FONT></B>"
    sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"
    simgz="<img src=Imagenes/botdelx.jpg id=imgz>"
    ccolor = "#f7eedd"
    cstilo = "border-style:outset;border-width:1;border-Color:#dddddd;"
    cbordert = "border-top-style:ridge;border-top-width:1;border-top-Color:#dddddd;"
    cborderb = "border-bottom-style:ridge;border-bottom-width:1;border-bottom-Color:#dddddd;"
    cborderl = "border-left-style:ridge;border-left-width:1;border-left-Color:#dddddd;"
    cborderr = "border-right-style:ridge;border-right-width:1;border-right-Color:#dddddd;"

      Set cn = OFV.conectar(OFV.strconn2)

          sHTML = OFV.MenuHeader("","")
          shtml=shtml &  chr(13) & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>" & chr(13)
       sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.getElementById('x1').style.display='block';" & chr(34)
       sinput2 = sinput2  & " TITLE='VER USUARIOS' ><b>" & sImagen4si & "</b></a>"


            tabp = OFV.gentabla("", "aaa ", "", "90%", "", "0", "", "-1", "-2", "", "", "", "no")

''' style='border-style:ridge;border-width:2;border-Color:#f7aa22;'

            CELDA =  OFV.gencelda("", "VT STYLE='BORDER-STYLE:OUTSET;BACKGROUND-COLOR:#F7EEDD;BORDER-WIDTH:1;' ", "", "2%", "", "", "", sinput2, "SI")
            CELDA =  CELDA & OFV.gencelda("", "VT STYLE='BORDER-STYLE:OUTSET;BACKGROUND-COLOR:#F7EEDD;BORDER-WIDTH:1;' ", "", "", "", "", "","AREAS-DEPARTAMENTOS-SECTORES ", "SI")
            tabp = tabp & OFV.GenRow("", "", "", "", "", "", "", CELDA, "SI")
            tabp = tabp & OFV.GenRow("", "", "", "", "", "", "", "", "no")
            tabp = tabp & OFV.gencelda("", "AAA COLSPAN=2 ", "center", "", "", "", "", "", "no")

            tabp = tabp & OFV.gentabla("", "aaa ", "", "100%", "", "0", "", "-1", "-2", "#ffffff", "", "", "no")

      nID = ""
      nlevel = 0
      sSector = ""
      nSeg = ""
      nID = Request.querystring("ID")
      sSector = Request.querystring("Sector")
      sSector = OFV.ConvertirCar(sSector, "_", " ")
      sCodin = "&nbsp;"

         Strsql = "select * from csOrgmaster where ecodepend is null"
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         do until rs.eof
            nlevelid1    = rs(0)
            nlevel      = 1
         rs.movenext
         loop


          sSQL = "Select descri,codinter From eco Where eco.eco = " & nlevelid1
          Set rs = OFV.crearconsultaEx(sSQL, cn, 1, "")
          If Not rs.EOF Then
              sSector = replace(rs(0)," ","&nbsp;")
              sCodin = rs(1)
          End If
          Call OFV.cerrarconsulta(rs)


           celda = OFV.gencelda("", "aaa colspan=22", "", "", "", "", "", "&nbsp;", "si")
           fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

           celda = OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & cstilo & "'", "center", "", "", "", "",simagenE0, "si")


                slink = "<font face=tahoma color=#000080 size=1><b>" & sSector & "</b></font>"

           celda = celda & OFV.gencelda("", "aaa colspan=19 rowspan=2 ", "", "", "", "", "","&nbsp;" & slink, "si")
           celda = celda & OFV.gencelda("", "aaa ", "", "", "", "", "", "&nbsp;", "si")
           fila  = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

           celda = OFV.gencelda("", "aaa ", "", "", "", "", "", "&nbsp;", "si")
           fila  = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")


           tabp = tabp & fila

     nlevelid = "0" & nid
     if nlevelid > 0 then

       warrx = nid
       for xyy=1 to 10
          Strsql = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1 where ECO.eco = ECOREL.ECODEPEND"
          Strsql = Strsql & " and ECOREL.ECOMASTER = ECO_1.eco and ECO.eco = " & nlevelid
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if not rs.eof then
            warrx     = rs(0) & "#" & warrx
            nlevelid  = rs(0)


          else
            exit for
          end if

       next

       IF XYY = 1 THEN warrx  = nlevelid1 & "#" & warrx
       WARRD  = SPLIT(WARRX,"#")

       FOR XYY=1 TO UBOUND(WARRD)
       ''    RESPONSE.WRITE XYY & " " & WARRD(XYY) & " " & UBOUND(WARRD)
            sSQL = "Select descri,codinter From eco Where eco.eco = " & WARRD(XYY)
            Set rs = OFV.crearconsultaEx(sSQL, cn, 1, "")
            If Not rs.EOF Then
               sSector = replace(rs(0)," ","&nbsp;")
               sCodin = rs(1)
            End If
            Call OFV.cerrarconsulta(rs)

            xin = 1
            ReDim Preserve vecsec(xin)
            ReDim Preserve vecsecd(xin)

          'Sector Principal


              vecsec(1)  = CLng(WARRD(XYY-1))

              IF vecsec(1)  = nlevelid1 THEN
                 vecsec(1)  = 0
              END IF

              vecsecd(1) = sSector

              tabp = tabp & GeneraResto(vecsec, vecsecd,XYY + 1)

           nlevel = XYY + 1


       NEXT


     else

       nid = nlevelid1

            xin = 1
            ReDim Preserve vecsec(xin)
            ReDim Preserve vecsecd(xin)

            vecsec(1)  = 0
            vecsecd(1) = sSector


     end if

         ''   RESPONSE.WRITE UBOUND(vecsec)

            xin = 0
            vecsec(0)  = vecsec(1)
            vecsecd(0) = vecsecd(1)

            ReDim Preserve vecsec(xin)
            ReDim Preserve vecsecd(xin)

       sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.getElementById('x1').style.display='none';" & chr(34)
       sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
      ' sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     '  sinput3 = sinput3  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     '  sinput3 = sinput3  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.x1.style.display='none';" & chr(34)
      ' sinput3 = sinput3  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"



              sHTML = sHTML & "<DIV id='x1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:400px;top:10px;width:400px;height:280px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"
              'shtml=shtml & simgz
              sHTML = sHTML & "<SPAN><FONT FACE=VERDANA SIZE=1><b><P STYLE='border-style:outset;background-color:#eeeeee;' >"  &  sinput2 & "&nbsp;" & vecsecd(0) & "</P>"
              sHTML = sHTML & "</b></FONT>"
      Set cn2 = OFV.conectar(OFV.strconn0)

      strsql = "Select * From USUARIOS where SECTOR = " & NID & " ORDER BY DESCRIPCION "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then
           tabB = tabB & OFV.gentabla("", "aaa ", "", "100%", "", "0", "", "-1", "-2", "#ffffff", "", "", "no")

           DO UNTIL rsniv.eof
   ''         for xyy=1 to 10
             celda = OFV.gencelda("", "VT ", "CENTER", "2%", "", "", "",sImagen4si, "si")
             celda = celda & OFV.gencelda("", "VT ", "", "", "", "", "", "&nbsp;" & rsniv(1), "si")
             fila  = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
             tabB = tabB & fila
       ''     next
           rsniv.MOVENEXT
           LOOP
           tabB = tabB & OFV.genFtabla()
           sHTML = sHTML & tabB
        ELSE
              sHTML = sHTML & "<SPAN><FONT FACE=VERDANA SIZE=1><b><P >" & sinput2 & "&nbsp;NO HAY USUARIOS PARA EL SECTOR</P>"
              sHTML = sHTML & "</b></FONT>"
              sHTML = sHTML & "</SPAN>"
        END IF
        call ofv.cerrarconsulta(rsniv)





      Call OFV.cerrarconn(cn2)

              sHTML = sHTML & "</DIV>"




          sSQL = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1, segmento "
          sSQL = sSQL & "  where ECO.eco = ECOREL.ECOMASTER and "
          sSQL = sSQL & " ECOREL.ECODEPEND = ECO_1.eco and "
          sSQL = sSQL & " eco_1.ecoseg = segmento.codseg and "
          sSQL = sSQL & " ECO.eco = " & nID & " "


           Set rs = OFV.crearconsulta2(sSQL, cn)

           'Texto de Sectores Continuadores y Lineas
            Do Until rs.EOF
                xin = xin + 1
                ReDim Preserve vecsec(xin)
                ReDim Preserve vecsecd(xin)
                vecsec(xin) = CLng(rs(0))
                vecsecd(xin) = replace(rs(1)," ","&nbsp;")

                rs.movenext
            Loop
          Call OFV.cerrarconsulta(rs)




            tabp = tabp & GeneraResto(vecsec, vecsecd,nlevel + 1)

    celda = ""

    for xii=1 to 20
        celda = celda & OFV.gencelda("", "", "", "1%", "", "", "", "&nbsp;", "si")
    next

    celda = celda & OFV.gencelda("", "", "", "60%", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "", "", "2%", "", "", "", "&nbsp;", "si")

    fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

            tabp = tabp & fila
            tabp = tabp & OFV.genFtabla()




          sHTML = sHTML & "<center>" & tabp & "</center></font></center>"





      Call OFV.cerrarconn(cn)

      sHTML = sHTML & "</body></html>"


     Generar = sHTML







end if

Response.Write sHTML



%>