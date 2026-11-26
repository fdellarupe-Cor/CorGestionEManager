<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
Function GeneraStaff(vecsecS(), vecsecSD(), vecsec())

fila = ""

    bordert = ""
    borderl = ""
    borderr = ""
    bstilo = ""
    bcolor = ""
    ccolor = "#f7eedd"
    cstilo = "border-style:outset;border-width:1;border-Color:#dddddd;"
    cbordert = "border-top-style:ridge;border-top-width:1;border-top-Color:#cecece;"
    cborderl = "border-left-style:ridge;border-left-width:1;border-left-Color:#dddddd;"
    cborderr = "border-right-style:ridge;border-right-width:1;border-right-Color:#dddddd;"


For xinS = 1 To UBound(vecsecS) Step 2
    borderl = cborderl
    bordert = cbordert
    bstilo = cstilo
    celda = OFV.gencelda("", "aaa colspan=9 ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=9  style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

    celda = OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
    shref = "Organigrama.asp?ID=" & vecsecS(xinS) & "&Sector=" & OFV.ConvertirCar(vecsecSD(xinS), " ", "_")
    sinput = "<font face=tahoma color=#6699cc size=1>" & vecsecSD(xinS) & "</font>"
    slink = OFV.GenLink("", "a2", shref, "", "", "", sinput, "si")
    
    celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", ccolor, "", slink, "si")
    celda = celda & OFV.gencelda("", "aaa colspan=5 ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=5 style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
    If xinS < UBound(vecsecS) Then
        shref = "Organigrama.asp?ID=" & vecsecS(xinS + 1) & "&Sector=" & OFV.ConvertirCar(vecsecSD(xinS + 1), " ", "_")
        sinput = "<font face=tahoma color=#6699cc size=1>" & vecsecSD(xinS + 1) & "</font>"
        slink = OFV.GenLink("", "a2", shref, "", "", "", sinput, "si")
    
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

Function GeneraResto(vecsec(), vecsecd())

fila = ""


For xin = 1 To UBound(vecsec) Step 4

    bordert = ""
    borderl = ""
    borderr = ""
    bstilo = ""
    bcolor = ""
    ccolor = "#f7eedd"
    cstilo = "border-style:outset;border-width:1;border-Color:#dddddd;"
    cbordert = "border-top-style:ridge;border-top-width:1;border-top-Color:#dddddd;"
    cborderl = "border-left-style:ridge;border-left-width:1;border-left-Color:#dddddd;"
    cborderr = "border-right-style:ridge;border-right-width:1;border-right-Color:#dddddd;"

    borderl = cborderl
    celda = OFV.gencelda("", "aaa colspan=9 ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=9  style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
    borderl = ""

    
    vcant = UBound(vecsec) - xin + 1
    celda = OFV.gencelda("", "aaa colspan=3 ", "", "", "", "", "", "&nbsp;", "si")
    If vcant > 2 Then
       bordert = cbordert
       borderl = cborderl
       borderr = ""
    End If
    celda = celda & OFV.gencelda("", "aaa colspan=4  style='" & bordert & borderl & borderr & "'", "", "", "", "", "", "&nbsp;", "si")
    If vcant > 1 Then
       bordert = cbordert
       borderr = ""
       borderl = ""
       If vcant = 2 Or vcant > 3 Then borderl = cborderl
    End If
    celda = celda & OFV.gencelda("", "aaa colspan=2  style='" & bordert & borderl & borderr & "'", "", "", "", "", "", "&nbsp;", "si")
    borderr = ""
    If vcant = 1 Or vcant = 3 Or vcant > 4 Then
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
    If vcant = 2 Or vcant > 3 Then
       borderl = cborderl
    End If
       
    If vcant > 2 Then
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
    celda = OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
    If vcant > 2 Then
       bstilo = cstilo
       bcolor = ccolor
       shref = "Organigrama.asp?ID=" & vecsec(xin) & "&Sector=" & OFV.ConvertirCar(vecsecd(xin), " ", "_")
       sinput = "<font face=tahoma color=#000080 size=1>" & vecsecd(xin) & "</font>"
       slink = OFV.GenLink("", "a2", shref, "", "", "", sinput, "si")
       
    End If
    celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2  style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")
    celda = celda & OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
    
    If vcant = 2 Or vcant > 3 Then
       bstilo = cstilo
       bcolor = ccolor
       If vcant = 2 Then
          xinW = xin
       Else
          xinW = xin + 1
       End If

       shref = "Organigrama.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
       sinput = "<font face=tahoma color=#000080 size=1>" & vecsecd(xinW) & "</font>"
       slink = OFV.GenLink("", "a2", shref, "", "", "", sinput, "si")
    Else
       bstilo = ""
       bcolor = ""
       slink = "&nbsp;"
    End If
    celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")
    
    If vcant = 3 Or vcant = 1 Then
       bstilo = cstilo
       bcolor = ccolor
       If vcant = 3 Then
          xinW = xin + 1
       Else
          xinW = xin
       End If
       shref = "Organigrama.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
       sinput = "<font face=tahoma color=#000080 size=1>" & vecsecd(xinW) & "</font>"
       slink = OFV.GenLink("", "a2", shref, "", "", "", sinput, "si")
       celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")

    Else
       If vcant > 4 Then borderl = cborderl
       celda = celda & OFV.gencelda("", "aaa  rowspan=2 ", "", "", "", "", "", "&nbsp;", "si")
       celda = celda & OFV.gencelda("", "aaa  rowspan=2 style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
    End If
    

    If vcant = 2 Or vcant > 3 Then
       bstilo = cstilo
       bcolor = ccolor
       If vcant = 2 Then
          xinW = xin + 1
       Else
          xinW = xin + 2
       End If
       
       shref = "Organigrama.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
       sinput = "<font face=tahoma color=#000080 size=1>" & vecsecd(xinW) & "</font>"
       slink = OFV.GenLink("", "a2", shref, "", "", "", sinput, "si")
    Else
       bstilo = ""
       bcolor = ""
       slink = "&nbsp;"
    End If
    celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")
    celda = celda & OFV.gencelda("", "aaa colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
    
    If vcant > 2 Then
       bstilo = cstilo
       bcolor = ccolor
       If vcant > 3 Then
          xinW = xin + 3
       Else
          xinW = xin + 2
       End If
       shref = "Organigrama.asp?ID=" & vecsec(xinW) & "&Sector=" & OFV.ConvertirCar(vecsecd(xinW), " ", "_")
       sinput = "<font face=tahoma color=#000080 size=1>" & vecsecd(xinW) & "</font>"
       slink = OFV.GenLink("", "a2", shref, "", "", "", sinput, "si")
    Else
       bstilo = ""
       bcolor = ""
       slink = "&nbsp;"
    End If
    
    celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='" & bstilo & "'", "center", "18%", "", bcolor, "", slink, "si")
    celda = celda & OFV.gencelda("", "aaa colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

    celda = OFV.gencelda("", "aaa colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
    celda = celda & OFV.gencelda("", "aaa colspan=2  ", "", "", "", "", "", "&nbsp;", "si")
    fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")

  
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
    
      Set cn = OFV.conectar(OFV.strconn2)

      nID = ""

      sSector = ""
      nSeg = ""
      nID = Request.querystring("ID")
      sSector = Request.querystring("Sector")
      sSector = OFV.ConvertirCar(sSector, "_", " ")
      sCodin = "&nbsp;"
      If nID = "" And Request.Form.Count > 0 Then
          nID = Request.Form(1)
      End If
      If nID <> "" Then
          sSQL = "Select descri,codinter From eco Where eco.eco = " & nID
          Set rs = OFV.crearconsultaEx(sSQL, cn, 1, "")
          If Not rs.EOF Then
              sSector = rs(0)
              sCodin = rs(1)
          End If
          Call OFV.cerrarconsulta(rs)
        End If
        
      If nID <> "" Then
          sHTML = OFV.ReportHeader("Organigrama")


            tabp = OFV.gentabla("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "90%", "", "0", "", "-1", "-2", "", "", "", "no")
            tabp = tabp & OFV.GenRow("", "", "", "", "", "", "", "", "no")
            tabp = tabp & OFV.gencelda("", "", "center", "", "", "", "", "", "no")

            tabp = tabp & OFV.gentabla("", "", "", "100%", "", "0", "", "-1", "-2", "#ffffff", "", "", "no")

            xin = 0
            ReDim Preserve vecsec(xin)
            ReDim Preserve vecsecd(xin)

          'Sector Principal
          sSQL = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1 where ECO.eco = ECOREL.ECODEPEND"
          sSQL = sSQL & " and ECOREL.ECOMASTER = ECO_1.eco and ECO.eco = " & nID
           
           Set rs = OFV.crearconsulta2(sSQL, cn)
           If rs.EOF Then
              vecsec(xin) = 0
              vecsecd(xin) = sSector
           Else
              vecsec(xin) = CLng(rs(0))
              vecsecd(xin) = sSector
           End If
    
    
    
           celda = OFV.gencelda("", "aaa colspan=18", "", "", "", "", "", "&nbsp;", "si")
           fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
           
           celda = OFV.gencelda("", "aaa colspan=2", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2", "", "", "", "", "", "&nbsp;", "si")
           celda = celda & OFV.gencelda("", "aaa colspan=4 ", "", "", "", "", "", "&nbsp;", "si")
           If vecsec(0) <> 0 Then
                shref = "Organigrama.asp?ID=" & vecsec(0) & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")
                sinput = "<font face=tahoma color=#000080 size=1><b>" & vecsecd(0) & "</b></font>"
                slink = OFV.GenLink("", "a2", shref, "", "", "", sinput, "si")
           Else
                slink = "<font face=tahoma color=#bbbbbb size=1><b>" & vecsecd(0) & "</b></font>"
           End If
           
           celda = celda & OFV.gencelda("", "aaa colspan=2 rowspan=2 style='border-style:outset;border-width:1;border-color:#dddddd;'", "center", "18%", "", "#f7eedd", "", slink, "si")
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

          sSQL = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1, segmento "
          sSQL = sSQL & "  where ECO.eco = ECOREL.ECOMASTER and "
          sSQL = sSQL & " ECOREL.ECODEPEND = ECO_1.eco and "
          sSQL = sSQL & " eco_1.ecoseg = segmento.codseg and "
          sSQL = sSQL & " ECO.eco = " & nID & " and segmento.descri <> 'STAFF'"
    

           Set rs = OFV.crearconsulta2(sSQL, cn)

           'Texto de Sectores Continuadores y Lineas
            Do Until rs.EOF
                xin = xin + 1
                ReDim Preserve vecsec(xin)
                ReDim Preserve vecsecd(xin)
                vecsec(xin) = CLng(rs(0))
                vecsecd(xin) = rs(1)
            
                rs.movenext
            Loop
          Call OFV.cerrarconsulta(rs)

  

            xinS = 0
            ReDim Preserve vecsecS(xinS)
            ReDim Preserve vecsecSD(xinS)

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
                vecsecS(xinS) = CLng(rs2(0))
                vecsecSD(xinS) = rs2(1)
              rs2.movenext
          Loop

            tabp = tabp & GeneraStaff(vecsecS, vecsecSD, vecsec)
            tabp = tabp & GeneraResto(vecsec, vecsecd)

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
            
          'Tabla principal
             Tabla = ""
          sft = "<font face=tahoma color=#000077 size=1>&nbsp;"
          sft2 = "<font face=tahoma color=#770000 size=1><b>&nbsp;"
          Tabla = Tabla & "<br><br><center><font face=tahoma color=#000077 size=1>"
          sinput = sft2 & "Area:"
          celda = OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "10%", "", "", "", sinput, "si")
          sinput = sft & sSector
          celda = celda & OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "70%", "", "", "", "", sinput, "si")
          sinput = sft2 & "Codigo:"
          celda = celda & OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "10%", "", "", "", sinput, "si")
          sinput = sft & sCodin
          celda = celda & OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "10%", "", "", "", sinput, "si")
          fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
          Tabla = Tabla & OFV.gentabla("", "", "", "90% border=1 style='border-style:ridge;border-width:1;border-Color:#ffffff;'", "", "0", "", "-1", "-2", "", "", fila, "si")

          sSQL = "SELECT * FROM CSECOSNJ where Eco =" & nID
          Set rs = OFV.crearconsulta2(sSQL, cn)
          If Not rs.EOF Then
              sinput = sft2 & "Segmento:"
              celda = OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "", "", "", "", sinput, "si")
              sinput = sft & rs(1)
              celda = celda & OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "", "", "", "", sinput, "si")
              sinput = sft2 & "Nivel:"
              celda = celda & OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "", "", "", "", sinput, "si")
              sinput = sft & rs(2)
              celda = celda & OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "", "", "", "", sinput, "si")
              sinput = sft2 & "Jerarquia:"
              celda = celda & OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "", "", "", "", sinput, "si")
              sinput = sft & rs(3)
              celda = celda & OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "", "", "", "", sinput, "si")
              fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
              Tabla = Tabla & "<br>" & OFV.gentabla("", "", "", "90% border=1 style='border-style:ridge;border-width:1;border-Color:#ffffff;'", "", "0", "", "-1", "-2", "", "", fila, "si")
          End If

          sSQL = "SELECT * FROM CSECOPER where Eco =" & nID
          Set rs = OFV.crearconsulta2(sSQL, cn)
          If Not rs.EOF Then
              sinput = sft2 & "Perfiles Asociados:"
              celda = OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "", "", "", "", sinput, "si")
              sinput = sft2 & "Dotacion"
              celda = celda & OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "", "", "", "", sinput, "si")
              fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
              Do Until rs.EOF
                  sinput = sft & rs(1)
                  celda = OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "", "", "", "", sinput, "si")
                  sinput = sft & rs(2)
                  celda = celda & OFV.gencelda("", "aaa style='border-style:ridge;border-width:1;border-Color:#dddddd;'", "", "", "", "", "", sinput, "si")
                  fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
                  rs.movenext
              Loop
              Tabla = Tabla & "<br>" & OFV.gentabla("", "", "", "90% border=1 style='border-style:ridge;border-width:1;border-Color:#ffffff;'", "", "0", "", "-1", "-2", "", "", fila, "si")
          End If

            tabp = tabp & Tabla & "<br><br>"
            tabp = tabp & OFV.GenFcelda()
            tabp = tabp & OFV.GenFRow()
            tabp = tabp & OFV.genFtabla()



          sHTML = sHTML & "<center>" & tabp & "</center></font></center>"


      Else
          sHTML = OFV.FormHeader("Organigrama") & "<br><br>"
             Dim celdam, filam, tablam

          'Encabezados
             celdam = OFV.gencelda("", "cc", "", "100%", "", "", "", "<b>GRAFICO</b>", "si")
             filam = OFV.GenRow("", "", "", "", "", "", "", celdam, "si")

            'Grafico del organigrama
            
          celda = OFV.gencelda("", "cc", "center", "100%", "", "", "", "<b>Area&nbsp;-&nbsp;Sector</b>", "si")
             fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
          sSQL = "Select * From eco Order By descri"
          sinput = OFV.GenerarCombox2(sSQL, cn, 0, 1, "Descri")
          celda = OFV.gencelda("", "", "center", "100%", "", "", "", sinput, "si")
          fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
          'Boton Aceptar
          sinput = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "")
          celda = OFV.gencelda("", "", "center", "100%", "", "", "", sinput, "si")
             fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
             Tabla = OFV.gentabla("", "", "", "80% border=0", "", "0", "0", "0", "0", "", "", fila, "si")
             Formulario = OFV.GenForm("Organigrama", "", "Organigrama.asp", "Grama", "post", "", Tabla, "si")
             
             celdam = OFV.gencelda("", "", "center", "100%", "", "", "", Formulario, "si")
             filam = filam & OFV.GenRow("", "", "", "", "", "", "", celdam, "si")
             celdam = OFV.gencelda("", "", "", "100%", "", "", "", "&nbsp;", "si")
             filam = filam & OFV.GenRow("", "", "", "", "", "", "", celdam, "si")
             
             celdam = OFV.gencelda("", "cc", "", "100%", "", "", "", "<b>LISTA&nbsp;DEL&nbsp;ORGANIGRAMA</b>", "si")
             filam = filam & OFV.GenRow("", "", "", "", "", "", "", celdam, "si")


            'LISTA del organigrama
            
            'Boton Aceptar
          sinput = OFV.GenerarInput("", "Listar", "submit", "", "10", "bt", "")
             celda = OFV.gencelda("", "", "center", "100%", "", "", "", sinput, "si")
             fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
             Tabla = OFV.gentabla("", "", "", "80% border=0", "", "0", "0", "0", "0", "", "", fila, "si")
             Formulario = OFV.GenForm("lisorg", "", "lisOrganigrama.asp", "_blank", "post", "", Tabla, "si")
             
             celdam = OFV.gencelda("", "", "center", "100%", "", "", "", Formulario, "si")
             filam = filam & OFV.GenRow("", "", "", "", "", "", "", celdam, "si")
             celdam = OFV.gencelda("", "", "", "100%", "", "", "", "&nbsp;", "si")
             filam = filam & OFV.GenRow("", "", "", "", "", "", "", celdam, "si")
             
             
          tablam = OFV.gentabla("", "", "", "100% border=0", "", "0", "0", "0", "0", "#eeeeee", "", filam, "si")
             celdam = OFV.gencelda("", "", "", "100%", "", "", "", tablam, "si")
             filam = OFV.GenRow("", "", "", "", "", "", "", celdam, "si")
             tablam = OFV.gentabla("", "", "", "80% border=1", "", "0", "0", "0", "0", "", "", filam, "si")

          sHTML = sHTML & "<center>" & tablam & "</center>"

      End If

      Call OFV.cerrarconn(cn)

      sHTML = sHTML & "</body></html>"
    
    
     Generar = sHTML
    






end if

Response.Write sHTML



%>