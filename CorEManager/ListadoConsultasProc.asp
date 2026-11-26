<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

'''''''''''''''''''
     Generar = ""
    
     Addqry = ""
     Response.expires = 0

     Server.scripttimeout = 600
     Response.buffer = True

     'response.write request.querystring & "<br><br>"
     'response.write request.form

     XXXX    = NEWEMPRE("NO")
     smenu   = Request.querystring("Menu")
     sorden  = Request.querystring("sorden")
     sNombre = request.querystring("Nombre")
    
     If Len(sorden) = 0 Then
         sorden = "N"
     End If
    
     Sfiltro = Request.form("sfiltro")
    
        If Len(Sfiltro) = 0 Then
           Sfiltro = Request.querystring("sfiltro")
        End If

     If Len(Sfiltro) = 0 Then
         Sfiltro = "TODOS"
     End If

     Sfiltro = Replace(Sfiltro, "_", " ")
     CantRegMostrar = 10
     CantRegMover = CDbl(Request.querystring("CantRegMover"))
     ReDim Preserve bconsultar(4)
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "4"
     bconsultar(4) = OFV.AConsultar
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "2"
     bconsultar(2) = OFV.AConsultar
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "1"
     bconsultar(1) = OFV.AConsultar
     OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "3"
     bconsultar(3) = OFV.AConsultar


     sverestado = " and Objetos.objnivel >= " & Session("SegNivel") & " And Estados.objid In (9,"


     For i = 1 To 4
         If bconsultar(i) Then sverestado = sverestado & i & ","
     Next
     sverestado = Left(sverestado, Len(sverestado) - 1) & ") "

''''''     FILTRO POR EMPRESA

      addfrom =  ""





      if (SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1) or (SESSION("EMPMUN") = "N") THEN
         sverestado = sverestado & " and Objetos.OBJUN = " & Session("UN") & " "
      ELSE
         addfrom    =  " ,OBJMEMPALC "
         sverestado = sverestado & " and "
         sverestado = sverestado & " ( (OBJETOS.OBJUN = " & Session("UN") & ") OR "
         sverestado = sverestado & "   ( OBJETOS.OBJUN = 1 AND "
         sverestado = sverestado & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
         sverestado = sverestado & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
         sverestado = sverestado & " ) "
      END IF



         sverestado = sverestado & " and TIPOBJETO.TIPOCPT = 'D' "


'''''''''''
    If Request.Form.Count > 1 Then
       Robj = "F" 
    else
       Robj = "Q" 
    end if     


        


        If  Robj = "F" Then
        
            sID    = Request.Form("BUSCAR") & ""

            Addqry = "&BUSCAR=" & sID


        Else
            sID    = Request.querystring("BUSCAR") & ""

            Addqry = "&BUSCAR=" & sID

        End If
       


    Addqry = "&Menu=" & smenu & Addqry & "&Nombre=" & sNombre

    starget = "ListadoConsultasproc.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry
    sVolver = "ListadoConsultasproc.asp?Menu=" & smenu & sMenuAdd

    Addqry = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")

    Set cn = OFV.conectar(OFV.strconn2)


    procabuscar = "'NO'"

    strsql = "SELECT codproceso from csprocesos where "
    strsql = strsql & " (codproceso01 = '" & sID & "') or  "
    strsql = strsql & " (codproceso02 = '" & sID & "') or  "
    strsql = strsql & " (codproceso03 = '" & sID & "') or  "
    strsql = strsql & " (codproceso04 = '" & sID & "') or  "
    strsql = strsql & " (codproceso05 = '" & sID & "')     "
    strsql = strsql & " order by dscsecuen  "

    Set rsx = OFV.crearconsultaEx(strsql, cn, 1, "")
    do until rsx.eof

       procabuscar = procabuscar & ",'" & rsx(0) & "'"


    rsx.movenext
    loop
     
    procabuscar = "(" & procabuscar & ")"



    strsql = "SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, "
    strsql = strsql & "OBJETOS.codobjeto, OBJETOS.tipobjeto, Estados.estado as estdescri, "
    strsql = strsql & "OBJeven.descri, OBJeven.codproceso, Estados.estado as estdescri, "
    strsql = strsql & "OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, "
    strsql = strsql & "OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover, OBJETOS.OBJUN "



            strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, OBJEVEN, Estados " & addfrom
            strsql = strsql & " where "
            strsql = strsql & "OBJETOS.tipobjeto  = TIPOBJETO.tipobjeto and "
            strsql = strsql & "OBJPRTY.estado     = Estados.objid     and "
            strsql = strsql & "OBJETOS.codobjeto  = OBJPRTY.codobjeto and "
            strsql = strsql & "OBJETOS.codobjeto  = OBJEVEN.codobjeto and "
            strsql = strsql & "OBJEVEN.CODPROCESO in " & procabuscar & " "
            sTitulo = "PROCESO: " & sID

            sqle = "Select DSCPROC from CSPROCESOS where CODPROCESO =  '" & sID & "' "
            Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
            If Not rsx.EOF Then sTitulo = "PROCESO: " & RSX(0)
        

    strsql = strsql & sverestado

    sTitOrd  = ""
    sTitOrd2 = ""
 

        If Sfiltro <> "TODOS" Then
           sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
           strsql = strsql & " and tipobjeto.descri='" & Sfiltro & "' "
        End If
    
        Select Case sorden
                Case "N"
                        sTitOrd = sTitOrd & "Nombre"
                        strsql  = strsql & " order by objetos.descri,tipobjeto.descri "
                Case "T"
                        sTitOrd = sTitOrd & "Tipo"
                        strsql  = strsql & " order by tipobjeto.descri,objetos.descri "
                Case "C"
                        sTitOrd = sTitOrd & "Creacion"
                        strsql  = strsql & " order by objprty.fecemis,objetos.descri,tipobjeto.descri "
                Case "E"
                        sTitOrd = sTitOrd & "Estado"
                        strsql  = strsql & " order by estdescri,objetos.descri,tipobjeto.descri "
        End Select

'''   response.write "<br><br><br>" & strsql & "<br>"

    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

 ''   sHTML = OFV.MENUHeader("","")

 ''  response.write "<br><br><br><br>" & Request.form & "<br>"

    sHTML = ofv.FormHeader(session("FormN")) & chr(13)
    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs

  sImagenM = "<B><FONT FACE=""corflaticon"" COLOR=#005500 SIZE=3>2</FONT></B>"
  sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
  sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"
  sImagena = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=2>4</FONT></B>"



  Fila    = ""
  Celda   = ""
  sHref   = ""

  IF  Session("EMPMUN") = "S" THEN

  sLink   = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
  Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","","EMPRESA:&nbsp;&nbsp;","si")
  Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
  Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

  Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

  Tabla   = ofv.GenTabla("","aaa  ","","100% ","","0","","-1","-2","","",Fila,"si")
  sHTML   = sHTML & Tabla
  sHTML   = sHTML & "<br>"
  end if

  Fila = ""

    Celda  = OFV.GenCelda("", "ttt ", "", "", "", "", "", sImagena & "&nbsp;" & sTitulo & "", "si")
    Celda  = Celda & ofv.GenCelda("","ttt","right","30%","","","","&nbsp;" & sTitOrd2 & "Orden:&nbsp;" & sTitOrd & "&nbsp;&nbsp;","si")
    Fila   = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
    tablaz = OFV.GenTabla("", "", "", "100% ", "", "0", "", "-1", "0", "", "", Fila, "si")

   sHTML = sHTML & Tablaz

    Fila = ""


        
    Celda = OFV.GenCelda("", "ttr colspan=3", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=N target=_self ><B><u>&nbsp;ESQUEMA NORMATIVO&nbsp;</u></B></A>", "si")
    Celda = Celda & OFV.GenCelda("", "ttr", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=T target=_self ><B><u>&nbsp;TIPO&nbsp;</u></B></A>", "si")
    Celda = Celda & OFV.GenCelda("", "ttr", "", "", "", "", "", "Proceso", "si")
    Celda = Celda & OFV.GenCelda("", "ttr", "", "", "", "", "", "Evento", "si")
 '   Celda = Celda & OFV.GenCelda("", "ttr", "", "", "", "", "", "Tema/Subtema", "si")

    Fila = fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

    Response.write sHTML
    sHTML = ""
    Response.flush


'''''''''''''''''''' FIN ENCABEZADOS ''''''''''''''''''''''''''''''''''''''

    nLinea = 0

    fi = CLng(CDate("01-01-1900"))
    PosReg = 0

    xTrColor = "#FFFFFF"

    On Error Resume Next
    If Not rs.EOF Then rs.Move (CantRegMover)
    Do While (Not rs.EOF) And (PosReg < CantRegMostrar)

        If Not IsNull(rs(0)) Then stipo = rs(0)
        sEstado = rs(8)

        If Not IsNull(rs(9)) Then nIden = rs(9)
        If Not IsNull(rs(10)) Then sobjeto = rs(10)


        If Not IsNull(rs(19)) Then sobjeto = sobjeto & " - " & rs(19)
        If Not IsNull(rs(20)) Then sobjeto = sobjeto & " " & rs(20)

        sImagen = "<font face='corflaticon' size=3 color='#993333'>4</font>"

        sActcon = "Consultar.asp?Codigo=" & rs("codobjeto") & "&Tipo=" & rs("tipobjeto")
        sActcon = sActcon & "&Perfil=" & rs("perfil") & "&Datos=" & OFV.convertircar(sMenuMas, "&", "*")
        sActcon = sActcon & "&Volver=" & sVolver
        sHref = "&nbsp;"
        sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)

        sLink = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen, "si")

        Celda = OFV.GenCelda("", "ut", "", "3%", "", "", "", sLink, "si")

        smasc = "&nbsp;"
        If  rs(21) = 1 and  SESSION("UN") <> 1 and  SESSION("EMPMUN") = "S" Then
            smasc = "&nbsp;<font color=#993333  ><font class=btic title='Documento Corporativo' >G</font></font>"
        End If

        sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, sobjeto, "si")
        sTipoDsc = stipo
        sTipoDsc = "<a class=nt href=" & starget & "&sorden=" & sorden & "&sfiltro=" & Replace(stipo, " ", "_") & " target=_self >" & sTipoDsc & "</A>"

        Celda = Celda & OFV.GenCelda("", "ut ", "", "25%", "", "", "", sLink, "si")
        Celda = Celda & OFV.GenCelda("", "ut ", "", "4%", "", "", "", smasc, "si")

        Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "", sTipoDsc, "si")


        Strsql = "Select dscproc From csprocesos "
        strsql = strsql & " where codproceso = '" & rs(7) & "' "
 
        sInprc = "----"
        set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if Not rsx1.EOF then
           sInprc = rsx1(0)
        end if
 
        dVigencia = "---"


        Celda = Celda & OFV.GenCelda("", "ut ", "", "25%", "", "", "", sInprc, "si")
        Celda = Celda & OFV.GenCelda("", "ut ", "", "25%", "", "", "", rs(6), "si")





        Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF

        PosReg = PosReg + 1
        rs.movenext
        nLinea = nLinea + 1
        Response.write sHTML
        sHTML = ""
        If nLinea > 20 Then
            nLinea = 0
            Response.flush
        End If
        
    Loop
    sHTML = sHTML & OFV.GenTabla("", "", "", "100%  style={border-style:solid;border-width:1;border-color:#eeeeee;} ", "", "0", "", "-1", "0", "", "", Fila, "si")
    

    sAccion = "consproc.asp?Menu=" & sMenu & "&Nombre=" & sNombre


    Celda   = ofv.BotonVolver(sAccion)     


    sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", Celda)


    sHTML = sHTML  & "</body></html>"

    Session("sperfil") = ""

    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

    Generar = sHTML

'''''''''''''''''''

  sHTML = Generar

 '' sHTML = oClc.Generar



end if

response.write sHTML
response.flush

function vertexto(nwtexto)

    nwtextoF = nwtexto & ""

    if  nwtextoF <> "" then

        nwtextoA = split(nwtextoF," ")

        nwsep    = ""
        nwtextoF = ""

        for iii2 = 0 to ubound(nwtextoA)

            if len(nwtextoA(iii2)) > 2 then

                nwtextoF =  nwtextoF & nwsep & nwtextoA(iii2) 
         
                nwsep = "%"
            end if
        next


    end if 
    


    vertexto = nwtextoF

end function

%>