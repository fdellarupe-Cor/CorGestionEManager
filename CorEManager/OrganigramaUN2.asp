<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


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

      nID = "" 
      sSector = ""
      nSeg = ""

      nID = Request.querystring("ID")

      sUN = request.querystring("UN")
      IF sUN <> "" THEN
         SESSION("SUN") = SUN
      ELSE
         SUN = SESSION("SUN")
      END IF








          sHTML = OFV.MenuHeader("","")
             Dim celdam, filam, tablam

          'Encabezados
             celdam = OFV.gencelda("", "VT STYLE='border-bottom-style:ridge;border-bottom-width:2;border-bottom-color:#f7aa22;'", "center", "100%", "", "", "", "<b><big>ORGANIGRAMA</big></b>", "si")
             filam = OFV.GenRow("", "", "", "", "", "", "", celdam, "si")

             celdam = OFV.gencelda("", "", "", "100%", "", "", "", "&nbsp;", "si")
             filam = filam & OFV.GenRow("", "", "", "", "", "", "", celdam, "si")

            'Grafico del organigrama

          celda = OFV.gencelda("", "VT", "center", "100%", "", "", "", "&nbsp;", "si")
          fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")

''''''''''''''''''''''''



      xtseg = clng("0" & request.form("segm"))
      xtagr = clng("0" & request.form("agrm"))



      Celda = ofv.GenCelda("","vt","","","","","","&nbsp;","si")
      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
'--------------------

      Celda  = ofv.GenCelda("","tt","","10%","","","","Site/Sucursal","si")
      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

      Strsql = "Select * From sites where siteun = " & sun & "  Order By descri"
      sInput = ofv.GenerarCombo("Organigrama1","segm",xtseg,Strsql,cn,0,1,"")
      Celda  = ofv.GenCelda("","","","","","","",sInput,"si")
      formx  =  ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "OrganigramaUn.asp"
      formx  =  ofv.GenForm("Organigrama1","",sAccion,"","post","",formx,"si")

      Fila = Fila & formx
'-----------------------------
      Celda = ofv.GenCelda("","","","","","","","&nbsp;<br>","si")
      Fila = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

'-----------------------------
      Celda =ofv.GenCelda("","tt","","10%","","","","Nivel","si")
      Fila = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      Strsql = "Select distinct niveljerar.codnivel, niveljerar.descri  From eco,niveljerar where eco.ecoun = " & sun & " and eco.econiv = niveljerar.codnivel and eco.ecosite = " & xtseg & " Order By niveljerar.descri"
   ''   response.write strsql & "<br>"
      sInput = ofv.GenerarCombo("Organigrama2","agrm",xtagr,Strsql,cn,0,1,"")
      sInput = sinput & ofv.GenerarInput("segm",xtseg,"hidden","Organigrama2","1","dt","")
      Celda = ofv.GenCelda("","","","","","","",sInput,"si")
      formx  =  ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "OrganigramaUn.asp"
      formx  =  ofv.GenForm("Organigrama2","",sAccion,"","post","",formx,"si")

      Fila = Fila & formx

'-----------------------------

      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      Fila = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

'-----------------------------

      Celda = ofv.GenCelda("","tt","","10%","","","","Area-Sector","si")
      Fila = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      Strsql = "Select * From eco where ecoun = " & sun & " and econiv = " & xtagr & " and ecosite = " & xtseg & " Order By descri"
      sInput = ofv.GenerarCombox3(strSQL, cn, 0,1, "ListaP")

      if xtseg = 0 then xtseg = "nsnc"
      if xtagr = 0 then xtagr = "nsnc"

      sInput = sinput & ofv.GenerarInput("ListaS",xtseg,"hidden","Organigrama","1","dt","")
      sInput = sinput & ofv.GenerarInput("ListaG",xtagr,"hidden","Organigrama","1","dt","")
      Celda = ofv.GenCelda("","","","","","","",sInput,"si")
      formx  =  ofv.GenRow("","","","","30","","",Celda,"si")

'-----------------------------

      Celda = ofv.GenCelda("","","","","","","","&nbsp;<br>","si")
      formx  = formx  & ofv.GenRow("","","","","30","","",Celda,"si")

'-----------------------------

      sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
      Celda = ofv.GenCelda("","","Center ","","","","",sInput,"si")
      formx  = formx  & ofv.GenRow("","","","","30","","",Celda,"si")

'-----------------------------

      sAccion = "OrganigramaUnGraf.asp"
      formx = ofv.GenForm("Organigrama","",sAccion,"","post","",formx,"si")

      Formulario = fila & formx


''''''''''''''''''''''''''''



             celdam = OFV.gencelda("", "", "center", "100%", "", "", "", Formulario, "si")
             filam = filam & OFV.GenRow("", "", "", "", "", "", "", celdam, "si")
             celdam = OFV.gencelda("", "", "", "100%", "", "", "", "&nbsp;", "si")
             filam = filam & OFV.GenRow("", "", "", "", "", "", "", celdam, "si")


             celdam = OFV.gencelda("", "", "", "100%", "", "", "", "&nbsp;", "si")
             filam = filam & OFV.GenRow("", "", "", "", "", "", "", celdam, "si")


             tablam = OFV.gentabla("", "", "", "100% border=0", "", "0", "0", "0", "0", "", "", filam, "si")
             celdam = OFV.gencelda("", "aaa valign=middle align=center ", "", "100%", "", "", "", tablam, "si")
             filam = OFV.GenRow("", "", "", "", "", "", "", celdam, "si")
             tablam = OFV.gentabla("", "aa style='filter:shadow(color=" & Session("APLSHW") & ",direction=125);' ", "", "40% ", "30%", "0", "0", "0", "0", "", "", filam, "si")

             sHTML = sHTML & "<center>" & tablam  & "</center>"



      Call OFV.cerrarconn(cn)

       sHTML = sHTML  & "</body></html>"


     Generar = sHTML







end if

Response.Write sHTML

function pp()
          sSQL = "Select * From eco where ecoun = " & sun & " Order By descri"
          sinput = OFV.GenerarCombox2(sSQL, cn, 0, 1, "Descri")
          celda = OFV.gencelda("", "", "center", "100%", "", "", "", sinput, "si")
          fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")



          'Boton Aceptar
             sinput = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "")
             celda = OFV.gencelda("", "", "center", "100%", "", "", "", sinput, "si")
             fila = fila & OFV.GenRow("", "", "", "", "", "", "", celda, "si")
             Tabla = OFV.gentabla("", "", "", "80% border=0", "", "0", "0", "0", "0", "", "", fila, "si")
             Formulario = OFV.GenForm("Organigrama", "", "OrganigramaUnGraf.asp", "Grama", "post", "", Tabla, "si")
end function

%>