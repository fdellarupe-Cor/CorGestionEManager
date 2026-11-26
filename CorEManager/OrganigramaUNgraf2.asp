<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
Function GeneraStaff()

fila = ""

    bordert = ""
    borderl = ""
    borderr = ""
    bstilo = ""
    bcolor = ""
    ccolor = "#ededed"
    cstilo = "border-style:solid;border-width:1;border-Color:#000000;"
    cbordert = "border-top-style:solid;border-top-width:1;border-top-Color:#000000;"
    cborderl = "border-left-style:solid;border-left-width:1;border-left-Color:#000000;"
    cborderr = "border-right-style:solid;border-right-width:1;border-right-Color:#000000;"


For xinS = 1 To UBound(vecsecS) Step 2
    borderl = cborderl
    bordert = cbordert
    bstilo = cstilo
    celda = OFV.gencelda("", "aaa colspan=9 ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=9  style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

    celda = OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
    shref = "organigramaUNGraf.asp?ID=" & vecsecS(xinS) & "&Sector=" & OFV.ConvertirCar(vecsecSD(xinS), " ", "_")
    sinput = "<font face=tahoma color=#6699cc size=1>" & vecsecSD(xinS) & "</font>"
    slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

    celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", ccolor, "", slink, "si")
    celda = celda & OFV.gencelda("", "aaa colspan=5 ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=5 style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
    If xinS < UBound(vecsecS) Then
        shref = "organigramaUNGraf.asp?ID=" & vecsecS(xinS + 1) & "&Sector=" & OFV.ConvertirCar(vecsecSD(xinS + 1), " ", "_")
        sinput = "<font face=tahoma color=#6699cc size=1>" & vecsecSD(xinS + 1) & "</font>"
        slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

        celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", ccolor, "", slink, "si")
    Else
        celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 ", "center", "18%", "", "", "", "&nbsp;", "si")
    End If
    celda = celda & OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

    bordert = cbordert
    borderl = cborderl
    celda = OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=5  style='" & bordert & "'", "", "", "", "", "", "&nbsp;", "si")
    bordert = ""
    borderl = ""
    If xinS < UBound(vecsecS) Then bordert = cbordert
    If (xinS + 1) < UBound(vecsecS) Or UBound(vecsec) > 0 Then borderl = cborderl
    celda = celda & OFV.gencelda("", "aaa colspan=5  style='" & bordert & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

Next

GeneraStaff = fila

End Function

Function GeneraResto()

fila = ""

    ccolor = "#ffffff"
    cstilo = "border-style:solid;border-width:1;border-Color:#000000;"
    cbordert = "border-top-style:solid;border-top-width:1;border-top-Color:#000000;"
    cborderl = "border-left-style:solid;border-left-width:1;border-left-Color:#000000;"
    cborderr = "border-right-style:solid;border-right-width:1;border-right-Color:#000000;"

xin = 1

vniv = vecsecn(xin)

do until xin > UBound(vecsec)

       celda = OFV.gencelda("", "aaa colspan=9 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' ", "", "", "", "", "", "&nbsp;", "si")
       celda = celda & OFV.gencelda("", "aaa colspan=9  style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;" & cborderl & "'", "", "", "", "", "", "&nbsp;", "si")
       fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

vnivd = 0

xnivT = "F"

for xinn = xin to UBound(vecsecn)

    if vniv <> vecsecn(xinn) then
       vniv = vecsecn(xinn)
       xnivT = "N"
       exit for
    end if
    vnivd = vnivd + 1
next


   
'''response.write "xin=" & xin & " cant en nivel=" & vnivd & "<br>"
    xinn = xin

    xpass = 0

For xin = xinn To (xinn + vnivd - 1) Step 4

    xinnp = 4 * xpass

    xpass = xpass + 1


    bordert = ""
    borderl = ""
    borderr = ""
    bstilo = ""
    bcolor = ""


    borderl = cborderl
    celda = OFV.gencelda("", "aaa colspan=9 ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=9  style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
    borderl = ""

    vcant = vnivd - xinnp
    vcantr = 5

    if xnivT = "F" then
       vcantr = vcant
    end if

    celda = OFV.gencelda("", "aaa colspan=3 ", "", "", "", "", "", "&nbsp;", "si")

    If vcantr > 2 Then
       bordert = cbordert
       borderl = cborderl
       borderr = ""
    End If

    celda = celda & OFV.gencelda("", "aaa colspan=4  style='" & bordert & borderl & borderr & "' ", "", "", "", "", "", "&nbsp;", "si")

    If vcantr > 1 Then
       bordert = cbordert
       borderr = ""
       borderl = ""
       If (vcantr = vcant) and (vcant = 2 Or vcant > 3) Then
          borderl = cborderl
       end if 

       If (vcantr <> vcant) and (vcant > 2) Then
          borderl = cborderl
       end if 

    End If

    celda = celda & OFV.gencelda("", "aaa colspan=2  style='" & bordert & borderl & borderr & "'", "", "", "", "", "", "&nbsp;", "si")

    borderr = ""
    If vcantr = 1 Or vcantr = 3 Or vcantr > 4 Then
       borderl = cborderl
    Else
       borderl = ""
    End If

    If vcant > 1 Then
       bordert = cbordert
    Else
       bordert = ""
    End If

    celda = celda & OFV.gencelda("", "aaa colspan=2  style='" & bordert & borderl & borderr & "'", "", "", "", "", "", "&nbsp;", "si")

       bordert = ""
       borderl = ""
       borderr = ""
       If (vcantr = vcant) and (vcant = 2 Or vcant > 3) or (vcantr <> vcant) and (vcant > 3 ) Then
 ''''''   If vcant = 2 Or vcant > 3 Then
          borderl = cborderl
       End If

    If ((vcantr = vcant) and (vcant > 2)) or ((vcantr <> vcant) and (vcant > 1)) Then
       bordert = cbordert
       borderr = cborderr
    End If

    celda = celda & OFV.gencelda("", "aaa colspan=4  style='" & bordert & borderl & borderr & "'", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=3  ", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

       bordert = ""
       borderl = ""
       borderr = ""
       bstilo = ""
       bcolor = ""
       slink = "&nbsp;"

    celda = OFV.gencelda("", "vt colspan=2 ", "", "", "", "", "", "&nbsp;", "si")

    If ((vcant = vcantr) and (vcant > 2)) or ((vcant <> vcantr)) Then
       bstilo = cstilo
       bcolor = ccolor
       shref = "organigramaUNGraf.asp?ID=" & vecsec(xin) & "&Sector=" & OFV.ConvertirCar(vecsecd(xin), " ", "_")
       sinput = "<font face=tahoma color=#000080 size=1>" & vecsecd(xin) & "</font>"
       slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
    End If

    celda = celda & OFV.gencelda("", "vt colspan=2 rowspan=2  style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")
    celda = celda & OFV.gencelda("", "vt colspan=2 ", "", "", "", "", "", "&nbsp;", "si")

    If ((vcant = vcantr) and (vcant = 2 Or vcant > 3)) or ((vcant <> vcantr) and (vcant > 2)) Then
       bstilo = cstilo
       bcolor = ccolor
       If vcant = 2 Then
          xinW = xin
       Else
          xinW = xin + 1
       End If



       shref = "organigramaUNGraf.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
       sinput = "<font face=tahoma color=#000080 size=1>" & vecsecd(xinW) & "</font>"
       slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
    Else
       bstilo = ""
       bcolor = ""
       slink = "&nbsp;"
    End If

    celda = celda & OFV.gencelda("", "vt colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")

    If ((vcant = vcantr) and (vcant = 3 Or vcant = 1))  Then

   '''''''''' If vcant = 3 Or vcant = 1 Then
       bstilo = cstilo
       bcolor = ccolor
       If vcant = 3 Then
          xinW = xin + 1
       Else
          xinW = xin
       End If
       shref = "organigramaUNGraf.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
       sinput = "<font face=tahoma color=#000080 size=1>" & vecsecd(xinW) & "</font>"
       slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
       celda = celda & OFV.gencelda("", "vt colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")

    Else
       If vcantr > 4 Then borderl = cborderl
       celda = celda & OFV.gencelda("", "vt  rowspan=2 ", "", "", "", "", "", "&nbsp;", "si")
       celda = celda & OFV.gencelda("", "vt  rowspan=2 style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
    End If


    If ((vcant = vcantr) and (vcant = 2 Or vcant > 3)) or ((vcant <> vcantr) and (vcant > 3)) Then

  ''''''  If vcant = 2 Or vcant > 3 Then
       bstilo = cstilo
       bcolor = ccolor
       If vcant > 2 then
          xinW = xin + 2
       Else
          xinW = xin + 1
       End If

       shref = "organigramaUNGraf.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
       sinput = "<font face=tahoma color=#000080 size=1>" & vecsecd(xinW) & "</font>"
       slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
    Else
       bstilo = ""
       bcolor = ""
       slink = "&nbsp;"
    End If

    celda = celda & OFV.gencelda("", "vt colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")
    celda = celda & OFV.gencelda("", "vt colspan=2 ", "", "", "", "", "", "&nbsp;", "si")

    If ((vcant = vcantr) and (vcant > 2)) or ((vcant <> vcantr) and (vcant > 1)) Then

 '''''''   If vcant > 2 Then
       bstilo = cstilo
       bcolor = ccolor
       If vcant > 3 Then
          xinW = xin + 3
       ElseIf vcant > 2 then
          xinW = xin + 2
       Else
          xinW = xin + 1
       End If
       shref = "organigramaUNGraf.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
       sinput = "<font face=tahoma color=#000080 size=1>" & vecsecd(xinW) & "</font>"
       slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
    Else
       bstilo = ""
       bcolor = ""
       slink = "&nbsp;"
    End If

    celda = celda & OFV.gencelda("", "vt colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")
    celda = celda & OFV.gencelda("", "vt colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

    celda = OFV.gencelda("", "vt colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "vt colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "vt colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "vt colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")


Next

    if xnivT = "Ff" then

       celda = OFV.gencelda("", "aaa colspan=9 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' ", "", "", "", "", "", "&nbsp;", "si")
       celda = celda & OFV.gencelda("", "aaa colspan=9  style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;" & cborderl & "'", "", "", "", "", "", "&nbsp;", "si")
       fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

    end if

    xin = xinn + vnivd 

 loop


GeneraResto = fila


End Function


ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
      Generar = ""
      'response.buffer=false
      Set cn = OFV.conectar(OFV.strconn2)
      Set cn2 = OFV.conectar(OFV.strconn0)
      Set cn3 = OFV.conectar(OFV.strconn0)



      if request.querystring("mseg")<>"" then
	    session("mseg")=request.querystring("mseg")
      end if

      nID = ""

      sSector = ""
      nSeg = ""

      nID = Request.querystring("ID")




      IF SESSION("MSEG") = "E" THEN

         sSQL = "SELECT sector,nivel FROM usuarios where id = '" & session("usuario") & "'"
	 Set rs = OFV.crearconsulta2(sSQL, cn2)
         If Not rs.EOF Then
            nidu  = rs(0)
            nivel = rs(1)
         end if
         Call OFV.cerrarconsulta(rs)

         If nID = "" then
            nID = nidu
         end if


      ELSE
        nivel=0

        sSector = Request.querystring("Sector")
        sSector = OFV.ConvertirCar(sSector, "_", " ")
        sCodin = "&nbsp;"
        If nID = "" And Request.Form.Count > 0 Then
           nID = Request.Form(1)
        End If
      END IF

      nid = "0" & REPLACE(nid,"nsnc","")

      if not nid > "0" then

          sHTML = OFV.ReportHeader("<FONT SIZE=2 COLOR=#6699CC><B>" & Session("Cliente") & "</B></FONT>") & ofv.Repbody("CORDPO - ORGANIGRAMA")
          shtml=shtml & "<br><br><br><br><br><center><p><b><font face=tahoma size=4 color=#993333>DEBE SELECCIONAR UN AREA O SECTOR<br>PARA GRAFICAR.</FONT></P>" & chr(13)
         sHTML = sHTML  & "<br></body></html>"
      else 


              sSECSITE  = 0
              sSECNIVEL = 0
              SNIVELA   = "S/D"

      If nID <> "" Then
          sSQL = "Select descri,codinter,ECONIV From eco Where eco.eco = " & nID
          Set rs = OFV.crearconsultaEx(sSQL, cn, 1, "")
          If Not rs.EOF Then
              sSector = rs(0)
              sCodin = rs(1)
              sSECSITE  = Request.Form("SECSITE")
              sSECNIVEL = Request.Form("SECNIVEL")

              sSQL = "SELECT * FROM NIVELJERAR where CODNIVEL = 0" & RS(2)
              Set rs1 = OFV.crearconsulta2(sSQL, cn)
              If NOT rs1.EOF Then
                 SNIVELA      = RS1(1)
               '''  response.write "NIV " & SNIVELA
              END IF

          End If
          Call OFV.cerrarconsulta(rs)
        End If


        ' response.write NID & "----" & sSector
      sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=2>&#0128;</FONT></B>"
      sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=2>4</FONT></B>"
      sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"
      If nID <> "" Then
          sHTML = OFV.ReportHeader("<FONT SIZE=2 COLOR=#6699CC><B>" & Session("Cliente") & "</B></FONT>") & ofv.Repbody("CORDPO - ORGANIGRAMA")
          shtml=shtml &  chr(13) & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>" & chr(13)

            shtml=shtml & "<script language=javascript>" & chr(13)
            shtml=shtml & "function mostrarmenu(xusr,XUSRN){" & chr(13)
            shtml=shtml & "document.all.US1.innerHTML= XUSRN;" & chr(13)
            shtml=shtml & "document.all.MN1.href='DpoVerObjetivos.asp?Uid=' + xusr +'&ID=" & NID & "&sector="&replace(sSector," ","_") & "&cantregmover=0" &"';" & chr(13)
            shtml=shtml & "document.all.MN3.href='DpoConfPAccion.asp?Uid=' + xusr +'&ID=" & NID & "&sector="&replace(sSector," ","_") & "&cantregmover=0" &"';" & chr(13)
            shtml=shtml & "document.all.MN5.href='DpoVerEVAL.asp?Uid=' + xusr +'&ID=" & NID & "&sector="&replace(sSector," ","_") & "&cantregmover=0" &"';" & chr(13)
            shtml=shtml & "document.all.MN4.href='DpoVerCAL.asp?Uid=' + xusr +'&ID=" & NID & "&sector="&replace(sSector," ","_") & "&cantregmover=0" &"';" & chr(13)
            shtml=shtml & "document.all.x1.style.display='block';" & chr(13)
            shtml=shtml & "}" & chr(13)
            shtml=shtml & "</script>" & chr(13)
'----------------------------------------------
            sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
            sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
            sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.x1.style.display='none';" & chr(34)
            sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
            sHTML = sHTML & "<DIV id='x1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:200px;top:10px;width:400px;height:280px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"
            sHTML = sHTML & "<SPAN><FONT FACE=VERDANA SIZE=1><b><P STYLE='border-style:outset;background-color:#eeeeee;' >" & sinput2 & "&nbsp;OPCIONES</P>"
            sHTML = sHTML & "</b></FONT>"

            tabB = tabB & OFV.gentabla("", "aa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-1", "-2", "", "", "", "no")

            celda = OFV.gencelda("", "VT ", "CENTER", "5%", "", "", "",sImagenEM, "si")
            celda = celda & OFV.gencelda("US1", "VT ", "", "", "", "", "", "&nbsp;", "si")
            fila  = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

            celda = OFV.gencelda("", "VT COLSPAN=2 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ", "CENTER", "2%", "", "", "","&nbsp;", "si")
            fila  = FILA & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

            celda = OFV.gencelda("", "VT VALIGN=TOP ", "CENTER", "5%", "", "", "",sImagen4si, "si")

            sinput2 = "<a  id='MN1' class=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
  	    sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
            sinput2 = sinput2  & "none" & chr(34) & ";' href=xxx"
            sinput2 = sinput2  & "  ><b>ADMINISTRACION DE OBJETIVOS<BR>Y PLAN DE ACCION</b></a>"

            celda = celda & OFV.gencelda("", "VT ", "", "", "", "", "", "&nbsp;" & sinput2, "si")
            fila  = FILA & OFV.GenRow("", "", "", "", "30", "", "", celda, "si")

            celda = OFV.gencelda("", "VT VALIGN=TOP ", "CENTER", "5%", "", "", "",sImagen4si, "si")
	    sinput2 = "<a  id='MN2' class=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & "none" & chr(34) & ";' href=xxx"
	    sinput2 = sinput2  & "  ><b>ADMINISTRACION DEL PLAN DE ACCION</b></a>"

            celda = celda & OFV.gencelda("", "VT ", "", "", "", "", "", "&nbsp;" & sinput2, "si")
          ''fila  = FILA & OFV.GenRow("", "", "", "", "30", "", "", celda, "si")



            celda = OFV.gencelda("", "VT VALIGN=TOP ", "CENTER", "5%", "", "", "",sImagen4si, "si")
            sinput2 = "<a  id='MN3' class=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & "none" & chr(34) & ";' href=xxx"
	    sinput2 = sinput2  & "  ><b>CONFIRMACION DEL PLAN DE ACCION</b></a>"

	    celda = celda & OFV.gencelda("", "VT ", "", "", "", "", "", "&nbsp;" & sinput2, "si")
	    fila  = FILA & OFV.GenRow("", "", "", "", "30", "", "", celda, "si")


            celda = OFV.gencelda("", "VT VALIGN=TOP ", "CENTER", "5%", "", "", "",sImagen4si, "si")
	    sinput2 = "<a  id='MN5' class=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & "none" & chr(34) & ";' href=xxx"
	    sinput2 = sinput2  & "  ><b>EVALUACION DE DESEMPEÑO</b></a>"

	    celda = celda & OFV.gencelda("", "VT ", "", "", "", "", "", "&nbsp;" & sinput2, "si")
	    fila  = FILA & OFV.GenRow("", "", "", "", "30", "", "", celda, "si")


            if SESSION("MSEG")="E" then
            visible =" display:none;"
            sImagen4si="&nbsp;"
            else
             visible=""
            end if

            celda = OFV.gencelda("", "VT VALIGN=TOP  ", "CENTER", "5%", "", "", "",sImagen4si, "si")
            sinput2 = "<a  id='MN4' class=CC style='"&visible & "text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
	        sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
	        sinput2 = sinput2  & "none" & visible & chr(34) & ";' href=xxx"
	        sinput2 = sinput2  & "  ><b>CALCULO INDIVIDUAL</b></a>"


	    celda = celda & OFV.gencelda("", "VT ", "", "", "", "", "", "&nbsp;" & sinput2, "si")
	    fila  = FILA & OFV.GenRow("", "", "", "", "30", "", "", celda, "si")


            tabB = tabB & fila
            tabB = tabB & OFV.genFtabla()
            sHTML = sHTML & tabB
            sHTML = sHTML & "</DIV>"

 '----------------------------------------------

            tabp = OFV.gentabla("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;filter:shadow(color=#dddddd,direction=125);'", "", "90%", "", "0", "", "-1", "-2", "", "", "", "no")
            tabp = tabp & OFV.GenRow("", "", "", "", "", "", "", "", "no")
            tabp = tabp & OFV.gencelda("", "", "center", "", "", "", "", "", "no")

            tabp = tabp & OFV.gentabla("", "", "", "100%  ", "", "0", "", "-1", "-2", "", "", "", "no")

            xin = 0
            ReDim Preserve vecsec(xin)
            ReDim Preserve vecsecd(xin)
            ReDim Preserve vecsecn(xin)

          'Sector Principal



          sSQL = "SELECT ECO_1.* FROM ECO E , ECOREL ER, ECO  ECO_1 where E.eco = ER.ECODEPEND"
          sSQL = sSQL & " and ER.ECOMASTER = ECO_1.eco and E.eco = " & nID
         '  response.write strsql
           Set rs = OFV.crearconsulta2(sSQL, cn)
           If rs.EOF Then
              vecsec(xin)  = 0
              vecsecd(xin) = sSector
              vecsecn(xin) = 0
           Else
              vecsec(xin)  = CLng(rs(0))
              vecsecd(xin) = sSector
              vecsecn(xin) = rs(5)

           End If

''response.write session("mseg") & " " & nid & " " & nidu & " " & (cstr(nid) = cstr(nidu))

           if session("mseg") = "E" and cstr(nid) = cstr(nidu) then  vecsec(xin) = 0



           celda = OFV.gencelda("", "aaa colspan=18", "", "", "", "", "", "&nbsp;", "si")
           fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

           celda = OFV.gencelda("", "aaa colspan=2", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa colspan=4 ", "", "", "", "", "", "&nbsp;", "si")
           If vecsec(0) <> 0 Then
                shref = "organigramaUNGraf.asp?ID=" & vecsec(0) & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")
                sinput = "<font face=tahoma color=#000080 size=1><b>" & vecsecd(0) & "</b></font>"
                slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
           Else
                slink = "<font face=tahoma color=#993333 size=1><b>" & vecsecd(0) & "</b></font>"
           End If

           celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='border-style:solid;border-width:1;border-color:#000000;'", "center", "18%", "", "#bbddff", "", slink, "si")
           celda = celda & OFV.gencelda("", "aaa colspan=4 ", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")

           fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

           celda = OFV.gencelda("", "aaa colspan=2", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa colspan=4 ", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa colspan=4 ", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")

           fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

           tabp = tabp & fila


          sSQL = "SELECT ECO_1.* FROM ECO E, ECOREL ER, ECO  ECO_1, segmento "
          sSQL = sSQL & "  where E.eco = ER.ECOMASTER and "
          sSQL = sSQL & " ER.ECODEPEND = ECO_1.eco and "
          sSQL = sSQL & " eco_1.ecoseg = segmento.codseg and "
          sSQL = sSQL & " E.eco = " & nID & " and segmento.descri <> 'STAFF' order by ECO_1.ECONIV"

           ' response.write ssql
           Set rs = OFV.crearconsulta2(sSQL, cn)

           'Texto de Sectores Continuadores y Lineas
            Do Until rs.EOF
                xin = xin + 1
                ReDim Preserve vecsec(xin)
                ReDim Preserve vecsecd(xin)
                ReDim Preserve vecsecn(xin)
                vecsec(xin) = CLng(rs(0))
                vecsecd(xin) = rs(1)
                vecsecn(xin) = rs(5)
                rs.movenext
            Loop
          Call OFV.cerrarconsulta(rs)



            xinS = 0
            ReDim Preserve vecsecS(xinS)
            ReDim Preserve vecsecSD(xinS)
            ReDim Preserve vecsecSN(xinS)

          sSQL = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1, segmento "
          sSQL = sSQL & "  where ECO.eco = ECOREL.ECOMASTER and "
          sSQL = sSQL & " ECOREL.ECODEPEND = ECO_1.eco and "
          sSQL = sSQL & " eco_1.ecoseg = segmento.codseg and "
          sSQL = sSQL & " ECO.eco = " & nID & " and segmento.descri = 'STAFF'"


          Set rs2 = OFV.crearconsulta2(sSQL, cn)
          Do Until rs2.EOF
                xinS = xinS + 1
                ReDim Preserve vecsecS(xinS)
                ReDim Preserve vecsecSD(xinS)
                ReDim Preserve vecsecSn(xinS)
                vecsecS(xinS) = CLng(rs2(0))
                vecsecSD(xinS) = rs2(1)
                vecsecSn(xinS) = rs2(5)
              rs2.movenext
          Loop

            if ubound(vecsecS) > 0 then
               tabp = tabp & GeneraStaff()
            end if

            if ubound(vecsec) > 0 then
               tabp = tabp & GeneraResto()
            end if

			celda = OFV.gencelda("", "", "", "1.22%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "1.22%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "9%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "9%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "1.22%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "1.22%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "9%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "9%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "9%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "9%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "9%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "9%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "1.22%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "1.22%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "9%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "9%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "1.22%", "", "", "", "&nbsp;", "si")
			celda = celda & OFV.gencelda("", "", "", "1.22%", "", "", "", "&nbsp;", "si")
			fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

            tabp = tabp & fila
            tabp = tabp & OFV.genFtabla()

               sft = "<font face=tahoma color=#000077 size=1>&nbsp;"
			   sft2 = "<font face=tahoma color=#770000 size=1><b>&nbsp;"
			   sft3 = "<font face=tahoma color=#ff0000 size=1>&nbsp;"
               sft1 = "<font face=tahoma color=#770000 size=1>&nbsp;"

          '  response.write session("mseg")

              fila=""

          'Tabla principal


          'Tabla = sinput1 & sinput2

          Tabla = Tabla & "<br><br><center><font face=tahoma color=#000077 size=1>"
          sinput = sft2 & "AREA:"
          celda = OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
          sinput = sft & sSector
          celda = celda & OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "70%", "", "", "", "", sinput, "si")
          sinput = sft2 & "CODIGO:"
          celda = celda & OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
          sinput = sft & sCodin
          celda = celda & OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
          fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

          sinput = sft2 & "NIVEL:"
          celda = OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
          sinput = sft & sNIVELA
          celda = celda & OFV.gencelda("", "aaa COLSPAN=3 style='border-style:solid;border-width:1;border-Color:#336699;'", "70%", "", "", "", "", sinput, "si")

          fila = FILA & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
          Tabla = Tabla & OFV.gentabla("", "", "", "90% style='border-style:solid;border-width:1;border-Color:#336699;filter:shadow(color=#dddddd,direction=125);' ", "", "0", "", "-1", "-2", "", "", fila, "si")


          sSQL = "SELECT * FROM usuarios where sector = " & nID & " AND SITE2 = " & sSECSITE & " order by nivel "
         ' response.write ssql
          dim nusuarios
          nusuarios=0
          Set rs = OFV.crearconsulta2(sSQL, cn2)
          If Not rs.EOF Then
              tipo="<b><font face=corflaticon size=2 color=#aa0000 >4</font></b>"
              sinput = sft2 & "PERSONAL"
              celda = OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "40%", "", "", "", sinput, "si")
           '''   sinput = sft2 & "NIVEL/PERFIL"
              sinput = sft2 & "PERFIL/FUNCION"
              celda = celda & OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "40%", "", "", "", sinput, "si")
              fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
              dim TXESTACC0,TXESTACC1,TXESTACC2
              TXESTACC0=0
              TXESTACC1=0
              TXESTACC2=0
              flag=false
              susuario=""
              Do Until rs.EOF

                  sinput = tipo & "&nbsp;" & sft & "<b>"  & rs(1) & "</b>"
                  if not flag then
                    capo=rs(1)
                    flag=true
                  end if
                  sinput2 = "<a   style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
		  sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
		  '''sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "mostrarmenu('"& rs(0) &"','" & rs(1) & "');" & chr(34)
		  sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' "
   	          sinput2 = sinput2  & "  ><b>" & sinput & "</b></a>"

                  celda = OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", sinput2, "si")
                  XPUESTO = RS("NIVEL")
                  sSQL = "SELECT * FROM NIVELCONFIDEN where NCCod = " & XPUESTO
                ' response.write ssql
                  Set rsX = OFV.crearconsulta2(sSQL, cn2)
                  If Not rsX.EOF Then
                     XPUESTO = RSX(1)
                  ELSE
                     XPUESTO = "---"
                  END IF

              ''' esp blp ''''''
                  XPUESTO = ""
              ''' fin esp blp ''''''


                  XPUESTOP = RS("PUESTO")
                  sSQL = "SELECT * FROM TIPOPERFIL where TIPOPERFIL = '" & XPUESTOP & "' "
                ' response.write ssql
                  Set rsX = OFV.crearconsulta2(sSQL, cn)
                  If Not rsX.EOF Then
                     XPUESTOP = "<B>" & RSX(1) & "</B>"
                  ELSE
                     XPUESTOP = "---"
                  END IF

                  'XPUESTO = RS(25)
            ' sSQL = "SELECT * FROM DPOPUESTOS where CODIGO = '" & RS(25) & "'"
            ' Set rsX = OFV.crearconsulta2(sSQL, cn)
            ' If Not rsX.EOF Then
            '    XPUESTO = RSX(3)
            ' END IF
                '''  sinput =  sft & "&nbsp;" & XPUESTO & "<BR>&nbsp;&nbsp;" & XPUESTOP

                  sinput =  sft & "&nbsp;" & XPUESTOP
                  celda = celda & OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", sinput, "si")

                  fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
                  rs.movenext
                  tipo="<b><font face=corflaticon size=2 color=#6699cc >=</font></b>"

              Loop

              'total
                tipo="<b><font face=corflaticon size=2 color=#aa0000 >4</font></b>"
		 sinput =  "&nbsp;"
                 celda =  OFV.gencelda("", "aaa colspan=2 bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "50%", "", "", "", sinput, "si")
                 fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
		 Tabla = Tabla & "<br>" & OFV.gentabla("", "", "", "95% style='border-style:solid;border-width:1;border-Color:#336699;filter:shadow(color=#dddddd,direction=125);' ", "", "0", "", "-1", "-2", "", "", fila, "si")

          End If

            tabp = tabp & Tabla & "<br><br>"
            tabp = tabp & OFV.GenFcelda()
            tabp = tabp & OFV.GenFRow()
            tabp = tabp & OFV.genFtabla()



          sHTML = sHTML & "<center>" & tabp & "</center></font></center>"


      Else


      End If

      Call OFV.cerrarconn(cn)
      Call OFV.cerrarconn(cn2)
      Call OFV.cerrarconn(cn3)
      if request.querystring("mseg")="M" then
       sHTML = sHTML  & "</body></html>"
      else
      ' sHTML = sHTML  & "<br><span style='{font-family:verdana;color:#993333;font-size:9pt;font-weight:600;}'>Status Plan de Accion</span>" & helpIconos(2) & "</body></html>"
      '   sHTML = sHTML  & "<br>" & helpIconos(5) & "</body></html>"
         sHTML = sHTML  & "<br></body></html>"
      end if

     Generar = sHTML

  end if





end if

Response.Write sHTML


%>