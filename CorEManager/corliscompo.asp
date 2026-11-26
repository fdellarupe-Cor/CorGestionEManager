<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
    
  
    Response.expires = 0

    Server.scripttimeout = 600
    Response.buffer = True
    smenu = Request.querystring("Menu")
    
   
    CantRegMostrar = CDbl(Request.querystring("CantRegMostrar"))
    CantRegMover = CDbl(Request.querystring("CantRegMover"))
    OrdenarPor = Request.querystring("OrdenarPor")

    response.write "recibe " & Request.querystring & "---" & Request.Form & "<br>"
    
    
    If Len(Trim(OrdenarPor)) = 0 Then
      OrdenarPor = " descri2 "
    End If
    
   OrdenarPor = UCase(OrdenarPor)
   sOrdenx = Trim(OrdenarPor)
   Select Case OrdenarPor
    Case "DESCRI"
     OrdenarPor = " order by complementos.descri"
    Case Else
     OrdenarPor = " order by " & OrdenarPor
   End Select
   
    
   

    
      If smenu = "PORTEXTO" Then

        If Request.Form.Count > 0 Then
          
          sID = Replace(Request.Form("buscar"), "'", "")
          sID = Replace(sID, Chr(34), "")
        Else
           sID = Request.querystring("buscar")
           sID = OFV.convertircar(sID, "_", " ")
        End If


        
       
    ElseIf smenu = "PORIDENTIFICACION" Then

        If Request.Form.Count > 0 Then
        
          sID = Request.Form("buscar")
          
        Else
           sID = Request.querystring("buscar")
        End If

    Else

        If Request.Form.Count > 0 Then
        
        sID = Request.Form("lista")
        
        Else
           sID = Request.querystring("buscar")
        End If

       
    
     End If
     sID = OFV.convertircar(sID, "_", " ")
   Set cn = OFV.conectar(OFV.strconn2)

    

    Select Case smenu
    Case "PORTIPO"

        strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri as descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND tipocomp='" & sID & "' " & OrdenarPor
        Response.write strsql & "<br>"
        sTitulo = "Componente: " & sID
         
    Case "PORTEXTO"
        sTitulo = "Texto: " & sID
       
        strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri as descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND Complementos.DESCRI  like '%" & sID & "%' " & OrdenarPor
        Response.write strsql & "<br>"
        sTitulo = "Texto: " & sID
    Case "PORIDENTIFICACION"

        strsql = "Select Complementos.*,ARCHIVO,TIPOCOMP.descri as descri2 From Complementos,TIPOCOMP Where TIPOCOMP=TIPO AND codiden ='" & sID & "'  " & OrdenarPor
        Response.write strsql & "<br>"
        sTitulo = "IDENTIFICACION " & sID
        
    Case "PORSISTEMA"
       strsql = "select Complementos.*,ARCHIVO,TIPOCOMP.descri as descri2 from Compprty,complementos,TIPOCOMP"
       strsql = strsql & " where sistorigen='" & sID & "' "
       strsql = strsql & " and Compprty.codcomp=complementos.codcomp "
       strsql = strsql & " and Compprty.tipocomp=complementos.tipocomp "
       strsql = strsql & " and complementos.tipocomp=TIPOCOMP.TIPO  " & OrdenarPor
        Response.write strsql & "<br>"
    End Select
 
    
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
    Celda = OFV.GenCelda("", "tt align=center ", "", "", "", "", "", "Tipo", "si")
    sLink = "<a href=ListadoConsultasComp.asp?menu=" & smenu & "&OrdenarPor=Descri&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Descripción</a>"
    Celda = Celda & OFV.GenCelda("", "tt  align=center ", "", "", "", "", "", sLink, "si")
    sLink = "<a href=ListadoConsultasComp.asp?menu=" & smenu & "&OrdenarPor=descri2&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Tipo&nbsp;Componente</a>"
    Celda = Celda & OFV.GenCelda("", "tt  align=center ", "", "", "", "", "", sLink, "si")
    sLink = "<a href=ListadoConsultasComp.asp?menu=" & smenu & "&OrdenarPor=codiden&CantRegMostrar=" & CantRegMostrar & "&CantRegMover=" & CantRegMover & "&buscar=" & sID & " >Identificación</a>"
    Celda = Celda & OFV.GenCelda("", "tt  align=center  ", "", "", "", "", "", sLink, "si")
    Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    Response.write sHTML
    sHTML = ""
    Response.flush
   nLinea = 0
   PosReg = 0
   On Error Resume Next
   If Not rs.EOF Then rs.Move (CantRegMover)
     
    Do While (Not rs.EOF) And (PosReg < CantRegMostrar)
    sArchivo = "Iconos/" & rs(5)
   ' sImagen = OFV.GenImage("", "", "", "", "Iconos\" & simage, "0", "")
    sHref = "editarimagen2.asp?id=" & rs(0)
    sImagen = OFV.GenImage("", "", "", "", sArchivo, "0", "")
    sLink = OFV.GenLink("", "nt", sHref, "_blank", "", "", sImagen, "si")
    
    Celda = OFV.GenCelda("", "", "", "", "", "", "", sLink, "si")

    Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", rs(3), "si")
    Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", rs(6), "si")
    Celda = Celda & OFV.GenCelda("", "tu style={font-size:10;}", "", "", "", "", "", rs(1), "si")
    Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
    
    PosReg = PosReg + 1
    rs.movenext
    nLinea = nLinea + 1
        
    If nLinea > 20 Then
            nLinea = 0
            Response.flush
    End If
        
    Loop
    sHTML = sHTML & OFV.GenTabla("", "", "", "100% border=1 style={fontsize:2}", "", "0", "", "0", "0", "", "", Fila, "si") & "</body></html>"
    
     Addqry = "&Menu=" & smenu & "&OrdenarPor=" & sOrdenx
    
     Addqry = Addqry & "&buscar=" & sID
     sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", "")


    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

      response.write sHTML
    response.flush
%>