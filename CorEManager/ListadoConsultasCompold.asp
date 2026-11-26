<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
    Response.expires = 0

    Server.scripttimeout = 600
    Response.buffer = True
    smenu = Request.querystring("Menu")


    CantRegMostrar = 12
    CantRegMover   = CDbl(Request.querystring("CantRegMover"))
    OrdenarPor     = Request.querystring("OrdenarPor")
    sNombre        = request.querystring("Nombre")

    sNombreV       = request.querystring("Nombre")
    smenuV         = Request.querystring("Menu")

    If  Len(Trim(OrdenarPor)) = 0 Then
        OrdenarPor  = " descri2 "
    End If

   OrdenarPor      = UCase(OrdenarPor)
   sOrdenx         = Trim(OrdenarPor)
   Select Case OrdenarPor
          Case "DESCRI"
               OrdenarPor = " order by complementos.descri"
          Case Else
               OrdenarPor = " order by " & OrdenarPor
   End  Select



   smenu = ucase(smenu)

   If smenu = "PORTEXTO" Then

        If Request.Form.Count > 0 Then
           sIDv = Request.Form("buscar")

           sID = Replace(Request.Form("buscar"), "'", "")
           sID = Replace(sID, Chr(34), "")
        Else
           sIDv = Request.querystring("buscar")
           sID = Request.querystring("buscar")
           sID = OFV.convertircar(sID, "_", " ")
        End If

    ElseIf smenu = "PORIDENTIFICACION" Then

        If Request.Form.Count > 0 Then

          sID = Request.Form("buscar")
          sIDv = Request.Form("buscar")
        Else
           sID = Request.querystring("buscar")
           sIDv = Request.querystring("buscar")
        End If

    Else

        If Request.Form.Count > 0 Then

           sID = Request.Form("lista")
           sIDv = Request.Form("lista")
        Else
           sID = Request.querystring("buscar")
           sIDv = Request.querystring("buscar")
        End If

    End If

    If sID = "" Then
       sIDv = Request.querystring("buscar")
       sID = Request.querystring("buscar")
       sID = OFV.convertircar(sID, "_", " ")
    End If


   sID    = OFV.convertircar(sID, "_", " ")
   Set cn = OFV.conectar(OFV.strconn2)


'Response.write smenu & "<br>"
    Select Case ucase(smenu)
    Case "PORTIPO"

        strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND tipocomp='" & sID & "' " & OrdenarPor
    '    Response.write strsql & "<br>"
        sTitulo = "Componente: " & sID

    Case "PORTEXTO"
        sTitulo = "Texto: " & sID

        strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND Complementos.DESCRI  like '%" & sID & "%' " & OrdenarPor
      '  Response.write strsql & "<br>"
        sTitulo = "Texto: " & sID
    Case "PORIDENTIFICACION"

        strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri  descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND codiden ='" & sID & "'  " & OrdenarPor
     '   Response.write strsql & "<br>"
        sTitulo = "IDENTIFICACION " & sID

    Case "PORSISTEMA"
       strsql = "select Complementos.*,ARCHIVO,TIPOCOMP.descri  descri2 from Compprty,complementos,TIPOCOMP"
       strsql = strsql & " where sistorigen='" & sID & "' "
       strsql = strsql & " and Compprty.codcomp=complementos.codcomp "
       strsql = strsql & " and Compprty.tipocomp=complementos.tipocomp "
       strsql = strsql & " and complementos.tipocomp=TIPOCOMP.TIPO  " & OrdenarPor
      ' Response.write strsql & "<br>"
    End Select

   ' Response.write strsql & "<br>"
    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

    sHTML = OFV.FormHeader(sTitulo)
    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs

    sID = OFV.convertircar(sID, " ", "_")
    Celda = OFV.GenCelda("", "ttr align=center ", "", "", "", "", "", "Tipo", "si")
    sLink = "<a href=ListadoConsultasComp.asp?menu=" & smenu & "&OrdenarPor=Descri&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Descripción</a>"
    Celda = Celda & OFV.GenCelda("", "ttr  align=center ", "", "", "", "", "", sLink, "si")
    sLink = "<a href=ListadoConsultasComp.asp?menu=" & smenu & "&OrdenarPor=descri2&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Tipo&nbsp;Componente</a>"
    Celda = Celda & OFV.GenCelda("", "ttr  align=center ", "", "", "", "", "", sLink, "si")
    sLink = "<a href=ListadoConsultasComp.asp?menu=" & smenu & "&OrdenarPor=codiden&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Identificación</a>"
    Celda = Celda & OFV.GenCelda("", "ttr  align=center  ", "", "", "", "", "", sLink, "si")
    Fila = OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

    Response.write sHTML
    sHTML = ""
    Response.flush
   nLinea = 0
   PosReg = 0
   On Error Resume Next

    xTrColor = "#FFFFFF"

   If Not rs.EOF Then rs.Move (CantRegMover)

    Do While (Not rs.EOF) And (PosReg < CantRegMostrar)

    sImagen = "<font face=corflaticon size=4 color=#993333>2</font>"

    sHref2       = "EditarImagen2.asp?ID=" & rs(0) 
    sHref        = "javascript:void(0);"
 
    sLink = OFV.GenLink("", "nt onclick=""AbrirCons('" & sHref2 & "');return false;"" ", sHref, "_blank", "", "", sImagen, "si")

    Celda = OFV.GenCelda("", "ttt", "center", "2%", "", "", "", sLink, "si")


    sLink = OFV.GenLink("", "ttt onclick=""AbrirCons('" & sHref2 & "');return false;"" ", sHref, "_blank", "", "", rs(3), "si")

    Celda = Celda & OFV.GenCelda("", "ttt", "", "", "", "", "", sLink, "si")
    Celda = Celda & OFV.GenCelda("", "ut", "", "", "", "", "", rs(6), "si")
    Celda = Celda & OFV.GenCelda("", "ut", "", "", "", "", "", rs(1), "si")

    Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")

    if xTrColor = "#FFFFFF" then
       xTrColor = "#F9F9F9" 
    ELSE
       xTrColor = "#FFFFFF" 
    END IF

    PosReg = PosReg + 1
    rs.movenext
    nLinea = nLinea + 1

    If nLinea > 20 Then
            nLinea = 0
            Response.flush
    End If

    Loop
    sHTML = sHTML & OFV.GenTabla("", "", "", "95% ", "", "0", "", "-1", "0", "", "", Fila, "si")

    sAccion = "SubconsultaComponentes.asp?Menu=" & sMenu & "&Nombre=" & sNombre
    Celda   = ofv.BotonVolver(sAccion) 

     Addqry = "&Menu=" & smenuv & "&OrdenarPor=" & sOrdenx

     Addqry = Addqry & "&buscar=" & sIDv & "&nombre=" & snombrev

     sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", Celda)


     sHTML = sHTML  & "</body></html>"

    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

      response.write sHTML
    response.flush

    sHTML = ""
end if

response.write sHTML


%>