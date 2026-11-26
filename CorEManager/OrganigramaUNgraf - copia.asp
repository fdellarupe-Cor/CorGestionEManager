<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

'Class cClassName
'  public nombre
'  public apellido
'End Class


'set xx=new cClassName
'xx.nombre="12222222222222222222"


'Set d=Server.CreateObject("Scripting.Dictionary")

'd.add "1",xx


'response.write d.Item("1").nombre
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
    sinput = "<font face=verdana color=#6699cc size=1>" & vecsecSD(xinS) & "</font>"
    slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

    celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", ccolor, "", slink, "si")
    celda = celda & OFV.gencelda("", "aaa colspan=5 ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=5 style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
    If xinS < UBound(vecsecS) Then
        shref = "organigramaUNGraf.asp?ID=" & vecsecS(xinS + 1) & "&Sector=" & OFV.ConvertirCar(vecsecSD(xinS + 1), " ", "_")
        sinput = "<font face=verdana color=#6699cc size=1>" & vecsecSD(xinS + 1) & "</font>"
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
       sinput = "<font face=verdana color=#000080 size=1>" & vecsecd(xin) & "</font>"
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
       sinput = "<font face=verdana color=#000080 size=1>" & vecsecd(xinW) & "</font>"
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
       sinput = "<font face=verdana color=#000080 size=1>" & vecsecd(xinW) & "</font>"
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
       sinput = "<font face=verdana color=#000080 size=1>" & vecsecd(xinW) & "</font>"
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
       sinput = "<font face=verdana color=#000080 size=1>" & vecsecd(xinW) & "</font>"
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

   '  sSQL = "SELECT sector,nivel FROM usuarios where id = '" & session("usuario") & "'"
	' Set rs = OFV.crearconsulta2(sSQL, cn2)
    '     If Not rs.EOF Then
            nidu  = session("SGU_idcostcenter") 'rs(0)
            nivel = session("SGU_idnivel") 'rs(1)
     '    end if
      '   Call OFV.cerrarconsulta(rs)

         xxx=getinfouser(session("usuario"))

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

       ''   sHTML = OFV.ReportHeader("<FONT SIZE=2 COLOR=#6699CC><B>" & Session("Cliente") & "</B></FONT>") & ofv.Repbody("CORDPO - ORGANIGRAMA")

          sHTML = OFV.MenuHeader("","")

          shtml=shtml & "<br><br><br><br><br><center><p><b><font face=verdana size=4 color=#993333>DEBE SELECCIONAR UN AREA O SECTOR<br>PARA GRAFICAR.</FONT></P>" & chr(13)
         sHTML = sHTML  & "<br></body></html>"
      else


              sSECSITE  = 1
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


              IF  sSECSITE <> "" THEN
                  SESSION("SSITE") = sSECSITE
              ELSE
                  sSECSITE = SESSION("SSITE")
              END IF

              IF sSECNIVEL  = "" THEN sSECNIVEL  = 0


              sSQL = "SELECT * FROM ECOSITE where ECO = 0" & nID & " AND ECOSITE = 0" & sSECSITE
              Set rs1 = OFV.crearconsulta2(sSQL, cn)
              If  rs1.EOF Then
                 sSQL = "SELECT * FROM ECOSITE where ECO = 0" & nID
                 Set rs1 = OFV.crearconsulta2(sSQL, cn)
                 If  NOT rs1.EOF Then
                     sSECSITE = RS1(2)
                 END IF
              END IF

                  sSUCA = ""
                  sSUCN = ""

              sSQL = "SELECT * FROM SITES where OBJID = 0" & sSECSITE
              Set rs1 = OFV.crearconsulta2(sSQL, cn)
              If  NOT rs1.EOF Then
                  sSUCA = RS1(1)
                  sSUCN = RS1(3) & "&nbsp;"
              END IF



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
        ''  sHTML = OFV.ReportHeader("<FONT SIZE=2 COLOR=#6699CC><B>" & Session("Cliente") & "</B></FONT>") & ofv.Repbody("CORDPO - ORGANIGRAMA")

          sHTML = OFV.MenuHeader("","")
      ''    shtml = replace(shtml,"<body ","<body leftmargin=8 ")
          shtml=shtml &  chr(13) & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>" & chr(13)

            shtml=shtml & "<script language=javascript>" & chr(13)
            shtml=shtml & "function mostrarmenu(xusr,XUSRN){" & chr(13)
        '    shtml=shtml & "document.all.US1.innerHTML= XUSRN;" & chr(13)
        '    shtml=shtml & "document.all.MN1.href='DpoVerObjetivos.asp?Uid=' + xusr +'&ID=" & NID & "&sector="&replace(sSector," ","_") & "&cantregmover=0" &"';" & chr(13)
        '    shtml=shtml & "document.all.MN3.href='DpoConfPAccion.asp?Uid=' + xusr +'&ID=" & NID & "&sector="&replace(sSector," ","_") & "&cantregmover=0" &"';" & chr(13)
        '    shtml=shtml & "document.all.MN5.href='DpoVerEVAL.asp?Uid=' + xusr +'&ID=" & NID & "&sector="&replace(sSector," ","_") & "&cantregmover=0" &"';" & chr(13)
        '    shtml=shtml & "document.all.MN4.href='DpoVerCAL.asp?Uid=' + xusr +'&ID=" & NID & "&sector="&replace(sSector," ","_") & "&cantregmover=0" &"';" & chr(13)
        '    shtml=shtml & "document.all.x1.style.display='block';" & chr(13)
            shtml=shtml & "}" & chr(13)
            shtml=shtml & "</script>" & chr(13)
'----------------------------------------------
        '    sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
        '    sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
        '    sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.x1.style.display='none';" & chr(34)
        '    sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
        '    sHTML = sHTML & "<DIV id='x1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:200px;top:10px;width:400px;height:280px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"
        '    sHTML = sHTML & "<SPAN><FONT FACE=VERDANA SIZE=1><b><P STYLE='border-style:outset;background-color:#eeeeee;' >" & sinput2 & "&nbsp;OPCIONES</P>"
        '    sHTML = sHTML & "</b></FONT>"

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
       '     sHTML = sHTML & tabB
       '     sHTML = sHTML & "</DIV>"

 '----------------------------------------------

            tabp = "<br><br><hr><br>" & OFV.gentabla("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "90%", "", "0", "", "5", "-2", "", "", "", "no")
            tabp = tabp & OFV.GenRow("", "", "", "", "", "", "", "", "no")
            tabp = tabp & OFV.gencelda("", "", "", "", "", "", "", "", "no")

            tabp = tabp & OFV.gentabla("", "", "", "100%  ", "", "0", "", "-1", "-2", "", "", "", "no")

            xin = 0
            ReDim Preserve vecsec(xin)
            ReDim Preserve vecsecd(xin)
            ReDim Preserve vecsecn(xin)

          'Sector Principal



          sSQL = "SELECT ECO_1.eco,ECO_1.descri,ECO_1.codinter,Eco_1.NODO,Eco_1.ECOSEG,Eco_1.ECONIV,Eco_1.ECOJERAR,Eco_1.ecoUN,Eco_1.ecoSite,ER.CODREL FROM ECO E , ECOREL ER, ECO  ECO_1 where E.eco = ER.ECODEPEND"
          sSQL = sSQL & " and ER.ECOMASTER = ECO_1.eco and E.eco = " & nID
         '  response.write strsql
           Set rs = OFV.crearconsulta2(sSQL, cn)
           If rs.EOF Then
              vecsec(xin)  = 0
              vecsecd(xin) = sSector
              vecsecn(xin) = 0
           Else
'''''''''''''''
'sSECSITE


              DO UNTIL RS.EOF

'''''''''''''''
                 SCODREL = RS(9)

                sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL
    '          response.write ssql
                Set rs2 = OFV.crearconsulta2(sSQL, cn)
                If NOT rs2.EOF Then
                   sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL & " AND SITE = " & sSECSITE
                '  response.write strsql
                   Set rs2 = OFV.crearconsulta2(sSQL, cn)
                   If NOT rs2.EOF Then
                      vecsec(xin)  = CLng(rs(0))
                      vecsecd(xin) = sSector
                      vecsecn(xin) = rs(5)
                      EXIT DO
                   END IF
                ELSE
                   If sSECSITE = 1 Then
                      vecsec(xin)  = CLng(rs(0))
                      vecsecd(xin) = sSector
                      vecsecn(xin) = rs(5)
                      EXIT DO
                   END IF
                END IF


''''''''''''
              RS.MOVENEXT
              LOOP
              If rs.EOF Then
                 vecsec(xin)  = 0
                 vecsecd(xin) = sSector
                 vecsecn(xin) = 0
              END IF

'''''''''''''

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
                sinput = "<font face=verdana color=#000080 size=1><b>subir</b></font>"
                slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

                celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='border-style:solid;border-width:1;border-color:#000000;'", "center", "18%", "", "#feeedd", "", slink, "si")
                celda = celda & OFV.gencelda("", "aaa colspan=4 ", "", "", "", "", "", "&nbsp;", "si")
                celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2", "", "", "", "", "", "&nbsp;", "si")
                celda = celda & OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")

                 fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

                 celda = OFV.gencelda("", "aaa colspan=2", "", "", "", "", "", "&nbsp;", "si")
                 celda = celda & OFV.gencelda("", "aaa colspan=4 ", "", "", "", "", "", "&nbsp;", "si")
                 celda = celda & OFV.gencelda("", "aaa colspan=4 ", "", "", "", "", "", "&nbsp;", "si")
                 celda = celda & OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")

                fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

                celda = OFV.gencelda("", "aaa colspan=2", "", "", "", "", "", "&nbsp;", "si")
               celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2", "", "", "", "", "", "&nbsp;", "si")
              celda = celda & OFV.gencelda("", "aaa colspan=4 ", "", "", "", "", "", "&nbsp;", "si")


                shrefold = "organigramaUNGraf.asp?ID=" & vecsec(0) & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")


             '   shref = "organigramaUNGraf.asp?ID=" & nid & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")

            '    shref = "organigramaUNFunc.asp?ID=" & nid & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")


                shref = "genDSCSECTORNV.asp?SECTOR=" & nID & "&NOMSECTOR=" & OFV.ConvertirCar(vecsecd(0), " ", "_")
                sinput = "<font face=verdana color=#000080 size=1><b>" & vecsecd(0) & "</b></font>"
                slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")


           Else
                slink = "<font face=verdana color=#993333 size=1><b>" & vecsecd(0) & "</b></font>"
           End If

           celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='border-style:solid;border-width:1;border-color:#000000;'", "center", "18%", "", "#feeedd", "", slink, "si")
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


          sSQL = "SELECT DISTINCT ECO_1.eco,ECO_1.descri,ECO_1.codinter,Eco_1.NODO,Eco_1.ECOSEG,Eco_1.ECONIV,Eco_1.ECOJERAR,Eco_1.ecoUN,Eco_1.ecoSite,ER.CODREL,ER.ECOMASTER,ER.ECODEPEND "
          sSQL = sSQL & " FROM ECO E, ECOREL ER, ECO  ECO_1, segmento, ECOSITE ES "
          sSQL = sSQL & " where E.eco = ER.ECOMASTER and "
          sSQL = sSQL & " ER.ECODEPEND = ECO_1.eco and "
          sSQL = sSQL & " eco_1.ecoseg = segmento.codseg and ES.ECO = ECO_1.ECO AND "
          sSQL = sSQL & " E.eco = " & nID & " AND (ES.ECOSITE = 0" & SESSION("SSITE")
          sSQL = sSQL & " OR ES.ECOSITE = 0" & sSECSITE & ") and "
          sSQL = sSQL & " segmento.descri <> 'STAFF' order by ECO_1.ECONIV"

       '     response.write ssql
           Set rs = OFV.crearconsulta2(sSQL, cn)

           'Texto de Sectores Continuadores y Lineas
            Do Until rs.EOF
''''''''''''''''
                 SCODREL = RS(9)
               '  smaster = RS(10)
               '  sdepend = RS(11)
               '  response.write " codrel: " &  SCODREL & " master: " & smaster & " depen: " & sdepend & "<br>"

                sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL
            ''   response.write sSQL
                Set rs2 = OFV.crearconsulta2(sSQL, cn)
                If NOT rs2.EOF Then
                   sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL & " AND (SITE = 0" & SESSION("SSITE") & " OR SITE = 0" & sSECSITE & ") "
             ''     response.write sSQL
                   Set rs2 = OFV.crearconsulta2(sSQL, cn)
                   svasec = "NO"
                   If NOT rs2.EOF Then
              ''   response.write " -codrel: " &  SCODREL & " site: " & rs2(2) & " site sel: " & SESSION("SSITE") &  "<br>"

                      if clng("0" & SESSION("SSITE")) = clng("0" & sSECSITE) then
                         svasec = "OK"
                      elseif clng("0" & SESSION("SSITE")) = clng("0" & rs2(2)) then
                         svasec = "OK"
                      else
                         sSQL = "SELECT ER.* FROM ECOREL ER,ECORELSITE ES where ER.CODREL = ES.CODREL" & " AND ES.SITE = 0" & SESSION("SSITE") & " and ER.ECODEPEND = 0" & rs(0)
                  ''     response.write sSQL
                         Set rs2 = OFV.crearconsulta2(sSQL, cn)
                         if rs2.eof then
                            svasec = "OK"
                         end if
                      end if

                      if svasec = "OK" then
                         xin = xin + 1
                         ReDim Preserve vecsec(xin)
                         ReDim Preserve vecsecd(xin)
                         ReDim Preserve vecsecn(xin)
                         vecsec(xin) = CLng(rs(0))
                         vecsecd(xin) = rs(1)
                         vecsecn(xin) = rs(5)
                      end if
                  '''    EXIT DO
                   END IF
                ELSE
                   If SESSION("SSITE") = 1 Then
                      xin = xin + 1
                      ReDim Preserve vecsec(xin)
                      ReDim Preserve vecsecd(xin)
                      ReDim Preserve vecsecn(xin)
                      vecsec(xin) = CLng(rs(0))
                      vecsecd(xin) = rs(1)
                      vecsecn(xin) = rs(5)
                  '''    EXIT DO
                   END IF
                END IF

'''''''''''''






            rs.movenext
            Loop
          Call OFV.cerrarconsulta(rs)



            xinS = 0
            ReDim Preserve vecsecS(xinS)
            ReDim Preserve vecsecSD(xinS)
            ReDim Preserve vecsecSN(xinS)

          sSQL = "SELECT DISTINCT ECO_1.eco,ECO_1.descri,ECO_1.codinter,Eco_1.NODO,Eco_1.ECOSEG,Eco_1.ECONIV,Eco_1.ECOJERAR,Eco_1.ecoUN,Eco_1.ecoSite,ECOREL.CODREL FROM ECO, ECOREL, ECO  ECO_1, segmento, ECOSITE ES "
          sSQL = sSQL & "  where ECO.eco = ECOREL.ECOMASTER and "
          sSQL = sSQL & " ECOREL.ECODEPEND = ECO_1.eco and "
          sSQL = sSQL & " eco_1.ecoseg = segmento.codseg and ES.ECO = ECO_1.ECO AND "
          sSQL = sSQL & " ECO.eco = " & nID & " AND (ES.ECOSITE = 0" & SESSION("SSITE")
          sSQL = sSQL & " OR ES.ECOSITE = 0" & sSECSITE & ") and "
          sSQL = sSQL & " segmento.descri = 'STAFF'"


          Set rs2 = OFV.crearconsulta2(sSQL, cn)
          Do Until rs2.EOF

''''''''''''''''
                 SCODREL = RS2(9)

                sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL
             '  response.write strsql
                Set rs3 = OFV.crearconsulta2(sSQL, cn)
                If NOT rs3.EOF Then
                   sSQL = "SELECT * FROM ECORELSITE where CODREL = 0" & SCODREL & " AND (SITE = 0" & SESSION("SSITE") & " OR SITE = 0" & sSECSITE & ") "
             ''     response.write sSQL
                   Set rs3 = OFV.crearconsulta2(sSQL, cn)
                   svasec = "NO"
                   If NOT rs3.EOF Then
                      if clng("0" & SESSION("SSITE")) = clng("0" & sSECSITE) then
                         svasec = "OK"
                      elseif clng("0" & SESSION("SSITE")) = clng("0" & rs3(2)) then
                         svasec = "OK"
                      else
                         sSQL = "SELECT ER.* FROM ECOREL ER,ECORELSITE ES where ER.CODREL = ES.CODREL" & " AND ES.SITE = 0" & SESSION("SSITE") & " and ER.ECODEPEND = 0" & rs2(0)
                  ''     response.write sSQL
                         Set rs3 = OFV.crearconsulta2(sSQL, cn)
                         if rs3.eof then
                            svasec = "OK"
                         end if
                      end if

                      if svasec = "OK" then
                         xinS = xinS + 1
                         ReDim Preserve vecsecS(xinS)
                         ReDim Preserve vecsecSD(xinS)
                         ReDim Preserve vecsecSn(xinS)
                         vecsecS(xinS) = CLng(rs2(0))
                         vecsecSD(xinS) = rs2(1)
                         vecsecSn(xinS) = rs2(5)
                    '''  EXIT DO
                     end if
                   END IF
                ELSE
                   If SESSION("SSITE") = 1 Then
                      xinS = xinS + 1
                      ReDim Preserve vecsecS(xinS)
                      ReDim Preserve vecsecSD(xinS)
                      ReDim Preserve vecsecSn(xinS)
                      vecsecS(xinS) = CLng(rs2(0))
                      vecsecSD(xinS) = rs2(1)
                      vecsecSn(xinS) = rs2(5)
                    '''  EXIT DO
                   END IF
                END IF

'''''''''''''




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

               sft  = "<font face=verdana color=#000077 size=1>&nbsp;"
	       sft2 = "<font face=verdana color=#770000 size=1><b>&nbsp;"
	       sft3 = "<font face=verdana color=#ff0000 size=1>&nbsp;"
               sft1 = "<font face=verdana color=#770000 size=1>&nbsp;"
               sff = "</font>"

          '  response.write session("mseg")

              fila=""

          'Tabla principal


          'Tabla = sinput1 & sinput2

         ''' Tabla = Tabla & "<center>"

          XFECODSC = "N"

          sSQL = "SELECT * FROM ECODESCRIPCION WHERE ECO = 0" & nID
          Set rs2 = OFV.crearconsulta2(sSQL, cn)
          IF NOT rs2.EOF THEN XFECODSC = "S"




          Tabla = Tabla & "<br><br><font face=verdana color=#000077 size=1>"


          sinput = sft2 & "SUCURSAL:" & sff
          celda = OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
          sinput = sft & sSUCN & sSUCA & sff

          IF XFECODSC = "N" THEN
             celda = celda & OFV.gencelda("", "aaa COLSPAN=3 style='border-style:solid;border-width:1;border-Color:#336699;'", "70%", "", "", "", "", sinput, "si")
          ELSE
             celda = celda & OFV.gencelda("", "aaa  style='border-style:solid;border-width:1;border-Color:#336699;'", "70%", "", "", "", "", sinput, "si")

             sAccion = "genDSCSECTOR.asp?SECTOR=" & nID

    	     sHref   = sAccion
      	     sInput  = "<U>" & "DESCRIPCION" & "</U>"
             sInput  = ofv.GenLink("","ut" & "  ",sHref,"_blank","","",sInput,"si")

             celda = celda & OFV.gencelda("", "VT COLSPAN=2 bgcolor=#FACC66 title='VER DESCRIPCION DEL PUESTO.'  style='cursor:pointer;COLOR:#FFFFFF;border-style:solid;border-width:1;border-Color:#336699;'", "CENTER", "10%", "", "", "", sinput, "si")
          END IF

          fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

          sinput = sft2 & "AREA:" & sff
          celda = OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
          sinput = sft & sSector & sff
          celda = celda & OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "70%", "", "", "", "", sinput, "si")
          sinput = sft2 & "CODIGO:" & sff
          celda = celda & OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
          sinput = sft & sCodin & sff
          celda = celda & OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
          fila = FILA & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

          sinput = sft2 & "NIVEL:" & sff
          celda = OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "10%", "", "", "", sinput, "si")
          sinput = sft & sNIVELA & sff
          celda = celda & OFV.gencelda("", "aaa COLSPAN=3 style='border-style:solid;border-width:1;border-Color:#336699;'", "70%", "", "", "", "", sinput, "si")
          fila = FILA & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

          Tabla = Tabla & OFV.gentabla("", "", "", "100% bgcolor=#ffffff style='border-style:solid;border-width:1;border-Color:#cccccc;' ", "", "0", "", "1", "1", "", "", fila, "si")


       '   sSQL = "SELECT * FROM usuarios where sector = " & nID & " AND SITE2 = " & sSECSITE & " order by nivel "
       '   response.write ssql
          dim nusuarios
          nusuarios=0
        '  Set rs = OFV.crearconsulta2(sSQL, cn2)
        '  response.write  nID & "--" &  sSECSITE
          call getUsrInSect(nID,sSECSITE)
          arruid=split(session("SGU_UIGuserid"),chr(9))

            xineco = 0
            Strsql = "Select * FROM ECOPERFIL,tipoperfil WHERE ECOPERFIL.codperfil=tipoperfil.tipoperfil and  codsector = 0" & nID
			              'response.write strsql
			set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)


			do while not rsx2.eof
			  ReDim Preserve Vecoperfil(3,xineco)
			  Vecoperfil(0,xineco)=rsx2(0)  ' objid
			  Vecoperfil(1,xineco)=rsx2(2) ' codperfil
			  Vecoperfil(2,xineco)=0
			  Vecoperfil(3,xineco)=rsx2("descri")

			  xineco=xineco+1
			  rsx2.movenext
			loop
            call ofv.cerrarconsulta(rsx2)
            sImageva    = "&nbsp;&nbsp;" & "<font face='corflaticon' size=2 color='#00aa00'><b>&#0252;</b></font>"
            if session("SGU_UIGuserid") <> "" then
              sinput = sft2 & "PERFIL/FUNCION" & sff
			  celda =  OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "40%", "", "", "", sinput, "si")

              tipo="<b><font face=corflaticon size=2 color=#aa0000 >4</font></b>"
              sinput = sft2 & "PERSONAL" & sff
              celda = celda & OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "40%", "", "", "", sinput, "si")
           '''   sinput = sft2 & "NIVEL/PERFIL" & sff


              fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
              dim TXESTACC0,TXESTACC1,TXESTACC2
              TXESTACC0=0
              TXESTACC1=0
              TXESTACC2=0
              flag=false
              susuario=""






             for Iid=0 to ubound(arruid)
        '      Do Until rs.EOF
                  celda=""
                  'sinput = tipo & "&nbsp;" & sft & "<b>"  & rs(1) & "</b>" & sff
                  usrx=getinfouser(arruid(Iid))
                  sinput = tipo & "&nbsp;" & sft & "<b>"  & session("SGU_nombre") & "</b>" & sff
                  if not flag then
                   ' capo=rs(1)
                    capo=session("SGU_nombre")
                    flag=true
                  end if
                  sinput2 = "<a   style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
		          sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
		         '''sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "mostrarmenu('"& rs(0) &"','" & rs(1) & "');" & chr(34)
		          sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' "
   	              sinput2 = sinput2  & "  ><b>" & sinput & "</b></a>"

   	              sinput2 = "<b>" & sinput & "</b>"

                  celdaz = OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", sinput2, "si")
                 ' XPUESTO = RS("NIVEL")
                  XPUESTO=session("SGU_idnivel")
                  sSQL = "SELECT * FROM NIVELCONFIDEN where NCCod = 0" & XPUESTO
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

                  sImageva    = "&nbsp;&nbsp;" & "<font face='corflaticon' size=2 color='#00aa00'><b>&#0252;</b></font>"


                  XPUESTOP = session("SGU_codigopuesto") 'RS("PUESTO")

                  for Np=0 to xineco-1
                     if cstr(Vecoperfil(1,np) &"")=cstr(XPUESTOP &"") then
                        Vecoperfil(2,np)=1

                      exit for
                     end if
                  next


               '  RESPONSE.WRITE "CC=" & CC

                 sSQL = "SELECT * FROM TIPOPERFIL where TIPOPERFIL = '" & XPUESTOP & "' "
                ' response.write ssql
                  Set rsX = OFV.crearconsulta2(sSQL, cn)
                  If Not rsX.EOF Then


                     XPUESTOP = "<B>" & rsX(1) & "</B>"


                    Strsql = "Select * FROM ECOPERFIL WHERE "
                    'Strsql = Strsql & " codsector = 0" & nID & " AND codperfil ='" & RS("PUESTO") & "' "
                       Strsql = Strsql & " codsector = 0" & nID & " AND codperfil ='" & session("SGU_codigopuesto") & "' "
                   'RESPONSE.WRITE STRSQL & "<BR>"
                    set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                    If Not rsX2.EOF Then

                      sAccion = "genPTOFUN.asp?PUESTO=" & rsx2(0)

                      sstilo2 = "  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
		              sstilo2 = sstilo2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
		              sstilo2 = sstilo2  & "none" & chr(34) & ";' "
'
      	              xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
      	              sHref   = sAccion
      	              sInput  = sft & "&nbsp;" & XPUESTOP & sff & sImageva
      	              Celdaobs  = ofv.GenCelda("","ut" & xsffun & sstilo2 & " style='cursor:pointer;border-style:solid;border-width:1;border-Color:#336699;' title='VER DESCRIPCION DEL PUESTO.' ","","","1","","",sInput,"si")
                      Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"_blank","","",Celdaobs,"si")

                     celda = celda    & Celdaobs & celdaz

                   else

                      sinput =  sft & "&nbsp;" & XPUESTOP & sff
                      celda = celda &  OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", sinput, "si") & celdaz

                    end if

                  ELSE
                    XPUESTOP = "---"

                    sinput =  sft & "&nbsp;" & XPUESTOP & sff
                    celda = celda &  OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", sinput, "si") & celdaz

                  END IF





                  fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
          '        rs.movenext
                  tipo="<b><font face=corflaticon size=2 color=#6699cc >=</font></b>"

         '     Loop
                next

                For Np=0 to xineco-1
				 if  Vecoperfil(2,np)=0  then
				 sAccion = "genPTOFUN.asp?PUESTO=" & Vecoperfil(0,NP)

				 sstilo2 = "  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
				 sstilo2 = sstilo2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
				 sstilo2 = sstilo2  & "none" & chr(34) & ";' "
'
				  xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
				  sHref   = sAccion
				  sInput  = sft & "&nbsp;<b>" & Vecoperfil(3,NP) &"</b>" & sff & sImageva
				  Celdaobs  = ofv.GenCelda("","ut" & xsffun & sstilo2 & " style='cursor:pointer;border-style:solid;border-width:1;border-Color:#336699;' title='VER DESCRIPCION DEL PUESTO.' ","","","1","","",sInput,"si")
				  Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"_blank","","",Celdaobs,"si")
                  celdaz = OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", "---", "si")
				  celda =  Celdaobs & celdaz

                  fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
				             '        rs.movenext
                  tipo="<b><font face=corflaticon size=2 color=#6699cc >=</font></b>"
				  end if
                 next
              'total
                tipo="<b><font face=corflaticon size=2 color=#aa0000 >4</font></b>"
		        sinput =  "&nbsp;"
                 celda =  OFV.gencelda("", "aaa colspan=2 bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "50%", "", "", "", sinput, "si")
               '  fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
		         Tabla = Tabla & "<br>" & OFV.gentabla("", "", "", "100%  bgcolor=#ffffff style='border-style:solid;border-width:1;border-Color:#cccccc;' ", "", "0", "", "1", "1", "", "", fila, "si")




          else

          if xineco > 0 then
           sinput = sft2 & "PERFIL/FUNCION" & sff
		   celda =  OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "40%", "", "", "", sinput, "si")

		   tipo="<b><font face=corflaticon size=2 color=#aa0000 >4</font></b>"
		   sinput = sft2 & "PERSONAL" & sff
		   celda = celda & OFV.gencelda("", "aaa bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "40%", "", "", "", sinput, "si")
		   fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

                 For Np=0 to xineco-1
				 if  Vecoperfil(2,np)=0  then
				  sAccion = "genPTOFUN.asp?PUESTO=" & Vecoperfil(0,NP)

				 sstilo2 = "  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
				 sstilo2 = sstilo2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
				 sstilo2 = sstilo2  & "none" & chr(34) & ";' "
'
				  xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
				  sHref   = sAccion
				  sInput  = sft & "&nbsp;<b>" & Vecoperfil(3,NP) &"</b>" & sff & sImageva
				  Celdaobs  = ofv.GenCelda("","ut" & xsffun & sstilo2 & " style='cursor:pointer;border-style:solid;border-width:1;border-Color:#336699;' title='VER DESCRIPCION DEL PUESTO.' ","","","1","","",sInput,"si")
				  Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"_blank","","",Celdaobs,"si")
                  celdaz = OFV.gencelda("", "aaa style='border-style:solid;border-width:1;border-Color:#336699;'", "", "", "", "", "", "---", "si")
				  celda =  Celdaobs & celdaz

                  fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
				             '        rs.movenext
                  tipo="<b><font face=corflaticon size=2 color=#6699cc >=</font></b>"
				  end if
                 next
              'total
                tipo="<b><font face=corflaticon size=2 color=#aa0000 >4</font></b>"
		        sinput =  "&nbsp;"
                 celda =  OFV.gencelda("", "aaa colspan=2 bgcolor=#eeeeee style='border-style:solid;border-width:1;border-Color:#336699;'", "", "50%", "", "", "", sinput, "si")
               '  fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
		         Tabla = Tabla & "<br>" & OFV.gentabla("", "", "", "100%  bgcolor=#ffffff style='border-style:solid;border-width:1;border-Color:#cccccc;' ", "", "0", "", "1", "1", "", "", fila, "si")




           end if

          End If




         ' tabp = tabp & Tabla & "<br><br>"
		  tabp = tabp & OFV.GenFcelda()
		  tabp = tabp & OFV.GenFRow()
		  tabp = tabp & OFV.genFtabla()

        ''  sHTML = sHTML & "<center>" & tabp & "</center></font></center>"

          sHTML = sHTML &  tabp & "</font>"


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