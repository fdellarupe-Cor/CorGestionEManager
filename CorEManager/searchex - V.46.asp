

<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->



<%

ok = OFV.CheckUsuario()
If ok <> "Y" Then
    uv = OFV.uv
    If uv = "NO" Then ok = OFV.ValidarUsuario
    shtml = ok
Else
    response.expires = 0
    CantRegMostrarx = 8
    response.buffer = True

 ' on error resume next
    vuelta = OFV.vueltaasp("../", CantRegMostrarx, CantRegMoverx)

    Set CN = OFV.conectar(OFV.strconn2)

    Dim atit
    anchopag = "98%"

    search_phrase = replace(Request.querystring("YIDER")& "",chr(34),"")
     search_phrase=replace(search_phrase,"'","")
    If search_phrase = "" Then
        search_phrase = replace(Request.Form("buscar") & "",chr(34),"")
        search_phrase=replace(search_phrase,"'","")
    End If

    tipobsq = Request.querystring("b")
      search_strings = Split(search_phrase , " ", -1, 1)

	        xsdelim=""
	        sresult=""
		    For Each search_string In search_strings

		        If  len(search_string & "") > 3 Then


		           sresult= sresult & xsdelim & search_string
		           xsdelim= " "

		        End If

		    Next

  '  search_phrase=sresult 'trim(search_string)
    Session("yider") = search_phrase

    filtroEsq = ""

    If tipobsq <> "" Then
        delim = ""
        DELIM2 = ""
        SESQ = ""
        For i = 1 To Request.Form("ESQUEMAS").Count

            filtroEsq = filtroEsq & delim & "tipo='" & Request.Form("ESQUEMAS").Item(i) & "'"
            SESQ = SESQ & DELIM2 & "'" & Request.Form("ESQUEMAS").Item(i) & "'"
            DELIM2 = ","
            delim = " or "
        Next


        delim = ""

        filtrost = ""
        For i = 1 To Request.Form("subtemas").Count

            filtrost = filtrost & delim & Request.Form("subtemas").Item(i)

            delim = ","

        Next

    End If
    'xxx=armaglos()


    xclase1 = "bietx"

    xclase2 = "bietxG"
    m_search_style = "font-family:Arial; font-size:12pt; color:#000000;font-weight: bold"
    m_search_style2 = "font-family:Arial; font-size:18pt; color:#000000;font-weight: bold"
    m_style_you_searched_for = "font-family:Arial; font-size:10pt; color:#000000"
    m_style_title = "font-size:9pt; font-family:Arial;color:#B43104;font-weight: bold;text-decoration: underline;font-style: italic"
    m_style_titlex = "font-size:15pt; font-family:Arial;color:#000000;font-weight: bold;text-decoration: underline;font-style: italic"
    m_style_text = "font-size:8pt; font-family:verdana;xfont-style: italic"
    m_style_noise_word = "font-size:10pt; font-family:Arial; color:#6f6f6f"
    m_style_url = "font-size:10pt; font-family:Arial; color:green"
    m_style_br = "font-size:10pt; font-family:Arial"
    m_style_more_results_text = "font-size:10pt; font-family:Arial;color:#000000;"
    m_style_more_results_link = "font-size:10pt; font-family:Arial;color:#000000;text-decoration: underline"
    m_style_more_results_link_Next = "font-size:12pt; font-family:Arial; font-weight:bold;color:#000000;text-decoration: underline"


    shtml = OFV.FormHeader(Session("FormN"))
    shtml = shtml & CorPnlFnc("N", "")
    shtml = shtml & "<script languaje=javascript src='incfile/tooltip.js'></script>"
    shtml = shtml & "<script languaje=javascript > var posp=1;var cantdoc=0;var topepag=0;pagant=1;"


    shtml = shtml & "function mostrarp(dato){" & Chr(13)
    '  shtml = shtml & " for(var id=1;id <= topepag ;id++){"
    shtml = shtml & " var o=document.getElementById('pag' + pagant);" & Chr(13)
    shtml = shtml & " o.style.display='none';"
    ';}"

    shtml = shtml & " var o=document.getElementById('pag' + dato);" & Chr(13)
    shtml = shtml & " o.style.display='';"
    shtml = shtml & " var o=document.getElementById('lnk' + dato);" & Chr(13)
    shtml = shtml & " o.style.fontSize='15pt';o.style.color='#000000';"
    shtml = shtml & " var o=document.getElementById('lnk' + pagant);" & Chr(13)
    shtml = shtml & " o.style.fontSize='10pt';pagant=dato;o.style.color='#B43104';"
    shtml = shtml & " var op=document.getElementById('infoP');" & Chr(13)
    shtml = shtml & " op.innerHTML='<h5>Pagina ' + dato + ' de ' + topepag + '  - Total Documentos encontrados : ' + cantdoc +'<h5>' ;" & Chr(13)



    shtml = shtml & "}"


    shtml = shtml & "function mostrarpav(dato){" & Chr(13)
	    '  shtml = shtml & " for(var id=1;id <= topepag ;id++){"
	    shtml = shtml & " var o=document.getElementById('pag' + pagant);" & Chr(13)
	    shtml = shtml & " o.style.display='none';"
	    ';}"

	    shtml = shtml & " var o=document.getElementById('pag' + dato);" & Chr(13)
	    shtml = shtml & " o.style.display='';"
	    shtml = shtml & " var o=document.getElementById('lnkav' + dato);" & Chr(13)
	    shtml = shtml & " o.style.fontSize='15pt';o.style.color='#000000';"
	    shtml = shtml & " var o=document.getElementById('lnkav' + pagant);" & Chr(13)
	    shtml = shtml & " o.style.fontSize='10pt';pagant=dato;o.style.color='#B43104';"
	    shtml = shtml & " var op=document.getElementById('infoPAV');" & Chr(13)
	    shtml = shtml & " op.innerHTML='<h5>Pagina ' + dato + ' de ' + topepag + '  - Total Documentos encontrados : ' + cantdoc +'<h5>' ;" & Chr(13)



    shtml = shtml & "}"



    shtml = shtml & "function mostrarpox(dato){" & Chr(13)
    shtml = shtml & " if (dato == pagant){ return; }"
    shtml = shtml & " var o=document.getElementById('pag' + dato);" & Chr(13)
    shtml = shtml & " o.style.display='';"
    shtml = shtml & " var o=document.getElementById('pag' + pagant);" & Chr(13)
    shtml = shtml & " o.style.display='none';pagant=dato;"
    shtml = shtml & " var op=document.getElementById('infoP');" & Chr(13)
    shtml = shtml & " op.innerHTML='<h5>Pagina ' + dato + ' de ' + topepag + '  - Total Documentos encontrados : ' + cantdoc +'<h5>' ;" & Chr(13)



    shtml = shtml & "}"

    shtml = shtml & "function mostrarelem(dato,dato2){" & Chr(13)

    shtml = shtml & " var o=document.getElementById(dato);" & Chr(13)
    shtml = shtml & " o.style.display='';"
    shtml = shtml & " var o=document.getElementById(dato2);" & Chr(13)
    shtml = shtml & " o.style.display='none';"
    shtml = shtml & "}"


    shtml = shtml & "function mostrar(dato){" & Chr(13)
    shtml = shtml & "if(dato=='S'){if (posp==topepag) {  return;} " & Chr(13)
    shtml = shtml & " var o=document.getElementById('pag' + posp);" & Chr(13)

    shtml = shtml & "o.style.display='none';posp++;pagant=posp;" & Chr(13)
    shtml = shtml & " if (posp==topepag) {document.getElementById('btsig').className='bietxG';}"

    shtml = shtml & "var o=document.getElementById('pag' + posp);" & Chr(13)
    shtml = shtml & " var op=document.getElementById('infoP');" & Chr(13)
    shtml = shtml & " op.innerHTML='<h5>Pagina ' + posp + ' de ' + topepag + '  - Total Documentos encontrados : ' + cantdoc +'<h5>';" & Chr(13)
    shtml = shtml & "document.getElementById('btant').className='bietx';" & Chr(13)

    shtml = shtml & "o.style.display='';}"
    shtml = shtml & "else{ if (posp==1) { return;}" & Chr(13)
    shtml = shtml & " var o=document.getElementById('pag' + posp);" & Chr(13)

    shtml = shtml & "o.style.display='none';posp--;pagant=posp;" & Chr(13)
    shtml = shtml & " if (posp==1) {document.getElementById('btant').className='bietxG';}"
    shtml = shtml & " if (posp < topepag) {document.getElementById('btsig').className='bietx';}"
    shtml = shtml & " var op=document.getElementById('infoP');" & Chr(13)
    shtml = shtml & " op.innerHTML='<p><h5>Pagina ' + posp + ' de ' + topepag + '  - Total Documentos encontrados : ' + cantdoc +'<h5></p>';" & Chr(13)
    shtml = shtml & "var o=document.getElementById('pag' + posp);" & Chr(13)
    shtml = shtml & "o.style.display='';"

    shtml = shtml & "}}"

    shtml = shtml & "</script>"


    sql = "select orden,b.valor1,a.descri from filtrobsq b  "
    sql = sql & " inner join tipobjeto a on(a.tipobjeto COLLATE DATABASE_DEFAULT =b.valor1 COLLATE DATABASE_DEFAULT )"

    sql = sql & " where b.tipo='ESQUEMA' ORDER BY b.ORDEN"

    Set rsESQ = OFV.crearconsultaEx(sql, CN, 1, volver)
    SESQUEMAS = "Esquemas normativos:  "

    delim = ""
    DELIM2 = ""
    orden_doc = ""
    nohay2=false
   if rsESQ.EOF then
    nohay2=true
    orden_doc = orden_doc & " WHEN ''  THEN 0"  & Chr(13)
   end if
    sinputesquemas = "<select multiple name=esquemas style='display:none'>"
    aesq = Split(SESQ, ",")

    Do While Not rsESQ.EOF

        orden_doc = orden_doc & " WHEN '" & rsESQ(1) & "'  THEN " & rsESQ(0) & Chr(13)
        If Request.Form("ESQUEMAS").Count = 0 Then
            filtroEsq = filtroEsq & delim & "tipo='" & rsESQ(1) & "'"
            SESQ = SESQ & DELIM2 & "'" & rsESQ(1) & "'"
            DELIM2 = ","
            delim = " or "
            sinputesquemas = sinputesquemas & " <option value=" & rsESQ(1) & " >" & rsESQ(2) & "</option>"

        Else
            For i = 1 To Request.Form("ESQUEMAS").Count

                If Request.Form("ESQUEMAS").Item(i) = rsESQ(1) Then
                    sinputesquemasZ = " <option value=" & rsESQ(1) & " selected >" & rsESQ(2) & "</option>"
                    Exit For
                Else
                    sinputesquemasZ = " <option value=" & rsESQ(1) & " >" & rsESQ(2) & "</option>"
                End If

            Next
            sinputesquemas = sinputesquemas & sinputesquemasZ
        End If


        rsESQ.movenext
    Loop
    sinputesquemas = sinputesquemas & "</select> "
    Call OFV.cerrarconsulta(rsESQ)
    sql = "select archivo,tipobjeto from tipobjeto  "


    Set rsESQ = OFV.crearconsultaEx(sql, CN, 1, volver)

    delim = ""
    orden_docFull = ""


    Do While Not rsESQ.EOF
        '  SESQUEMAS=SESQUEMAS & delim & rsESQ("descri")
        Iorden=trim(rsESQ(0) & "")
        if  Iorden =""  then
         Iorden="0"
        end if
        orden_docFull = orden_docFull & " WHEN '" & rsESQ(1) & "'  THEN " & Iorden & Chr(13)

        delim = " or "
        rsESQ.movenext
    Loop


    Call OFV.cerrarconsulta(rsESQ)
    sql = "select b.codprodserv ,  c.descri  +'/' + b.descri as descri2"

    sql = sql & "  from filtrobsq a"

    sql = sql & "  inner join     PRODSERV b on(a.valor1=b.codPRODSERV)"
    sql = sql & "  inner join     tema c  on( c.codtema=b.codtema)"
    sql = sql & "   where a.tipo='ST'"
    sql = sql & " ORDER BY ORDEN "



    Set rsESQ = OFV.crearconsultaEx(sql, CN, 1, volver)


    delim = ""
    ' filtrost=""
    if rsESQ.EOF then
    sinputSUBTEMAS=""
     nohay=true
    else
     nohay=false
     delim = ""
    sinputSUBTEMAS = "<select multiple name=SUBTEMAS >"
    Do While Not rsESQ.EOF
        If Request.Form("SUBTEMAS").Count = 0 Then
            filtrost = filtrost & delim & rsESQ(0)
            delim = ","

            sinputSUBTEMAS = sinputSUBTEMAS & " <option value=" & rsESQ(0) & ">" & rsESQ(1) & "</option>"
        Else
            For i = 1 To Request.Form("SUBTEMAS").Count

                If CStr(Request.Form("SUBTEMAS").Item(i) & "") = CStr(rsESQ(0) & "") Then
                    sinputesquemasZ = " <option value=" & rsESQ(0) & " selected >" & rsESQ(1) & "</option>"
                    Exit For
                Else
                    sinputesquemasZ = " <option value=" & rsESQ(0) & " >" & rsESQ(1) & "</option>"
                End If

            Next
            sinputSUBTEMAS = sinputSUBTEMAS & sinputesquemasZ
        End If


        rsESQ.movenext
    Loop
    sinputSUBTEMAS = sinputSUBTEMAS & "</select> <HR>"
    end if
    Call OFV.cerrarconsulta(rsESQ)

    If filtrost <> "" Then

        sql = "SELECT COBJ FROM CSRELPATHDESCRIEX WHERE MC7 IN(" & filtrost & ")"

        If SESQ <> "" Then

            sql = sql & "AND TOBJ IN(" & SESQ & ")"
        End If
        '  RESPONSE.WRITE SQL

        Set rsESQ = OFV.crearconsultaEx(sql, CN, 1, volver)
        delim = ""
        filtrost = ""
        sINDOC = False
        If rsESQ.EOF Then sINDOC = True
        Do While Not rsESQ.EOF

            filtrost = filtrost & delim & "CODIGO='" & rsESQ(0) & "'"
            delim = " OR "
            rsESQ.movenext
        Loop
        Call OFV.cerrarconsulta(rsESQ)
    End If

'response.write  sql


    If tipobsq = "" Then
        tablab = "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:auto;width:98%;border-style:solid;border-width:2;border-color:#ffffff;' >"
    Else
        tablab = "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:auto;width:98%;border-style:solid;border-width:2;border-color:#ffffff;display:none;' >"
    End If
    sAccion = "searchex.asp"
    sInputT = "<input type='text' class='dt' name='BUSCAR' id='textot' value=""" & search_phrase & """ size='100' style='height:20px;border-style:solid;border-width:1;border-color:#666666;' >"
    'BtnAva = "<input  id=btsig type=BUTTON  size=10%  value='Ir a Busqueda avanzada' class=" & xclase1 & "   onclick=""mostrarelem('DTBUSCARav','DTBUSCAR');""   >"
    if nohay and nohay2 then
	    BtnAva=""
	    else
	    BtnAva = "<input  id=btsig type=BUTTON  size=10%  value='Ir a Busqueda avanzada' class=" & xclase1 & "   onclick=""mostrarelem('DTBUSCARav','DTBUSCAR');""   >"
    end if
    sInput = OFV.GenerarInput("", "Buscar", "submit", "", "5", "at STYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' ", "")

    formx = OFV.GenForm("Subconsulta", "", sAccion, sTarget, "post", "", sInputT & sInput, "si")
    formx = formx & "<DIV ID=PAGINAS><h5>Calculando...<h5></DIV>"
    celda = OFV.GenCelda("", "vt valign=top", "", "", "", "", "", formx, "si")

    celda = celda & OFV.GenCelda("", "ss valign=top align=center", "", "70%", "", "", "", BtnAva & "<div style=" & m_search_style2 & "><b><p><div id=infoP ><h5>Calculando...<h5></div></div>", "si")

    siguiente = "<input  id=btsig type=BUTTON  size=10%  value='PAG. SIGUIENTE' class=" & xclase1 & "   onclick=""mostrar('S');""   >"
    ANTERIOR = "<input  id=btant type=BUTTON  size=10%  value='PAG. ANTERIOR' class=" & xclase2 & "   onclick=""mostrar('A');""   >"
    '  celda = celda & ofv.GenCelda("", "vt valign=top align=right", "", "70%", "", "", "", ANTERIOR & "&nbsp;&nbsp;&nbsp;" &siguiente, "si")
    'celda = celda & ofv.GenCelda("", "vt", "", "70%", "", "", "", siguiente, "si")
    sTarget = ""






    Fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
    ' celda=ofv.GenCelda("", "vt colspan=3 valign=top ", "", "", "", "", "",  "<DIV ID=PAGINAS><h5>Calculando...<h5></DIV>", "si")
    ' Fila = fila &  ofv.GenRow("", "", "", "", "", "", "", celda, "si")

    tablab = tablab & OFV.GenTabla("", "aaa     STYLE='border-style:solid;border-width:1;border-color:#ffcc00;' ", "", "", "", "0", "", "0", "0", "", "", Fila, "si")

    tablab = tablab & "<br>" & "</DIV>" & Chr(13)
    If tipobsq = "" Then
        tablabavan = "<DIV id='DTBUSCARav' STYLE='float: left; z-index:1;overflow:auto;width:98%;border-style:solid;border-width:2;border-color:#ffffff;display:none;' >"
    Else
        tablabavan = "<DIV id='DTBUSCARav' STYLE='float: left; z-index:1;overflow:auto;width:98%;border-style:solid;border-width:2;border-color:#ffffff;' >"
    End If
    sAccion = "searchex.asp?B=BA"
    sInputT = "<input type='text' class='dt' name='BUSCAR' id='textot' value=""" & search_phrase & """ size='100' style='height:20px;border-style:solid;border-width:1;border-color:#666666;' >"
    BtnAva = "<input  id=btsig type=BUTTON  size=10%  value='Ir a Busqueda Simple' class=" & xclase1 & "   onclick=""mostrarelem('DTBUSCAR','DTBUSCARav');""   >"
    sInput = OFV.GenerarInput("", "Buscar", "submit", "", "5", "at STYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' ", "")

    formx = OFV.GenForm("Subconsulta", "", sAccion, sTarget, "post", "", sInputT & sInput & "<br>" & sinputesquemas & sinputSUBTEMAS , "si")
    formx = formx & "<DIV ID=PAGINASAV><h5><h5></DIV>"
    celda = OFV.GenCelda("", "vt valign=top", "", "", "", "", "", formx, "si")

    celda = celda & OFV.GenCelda("", "ss valign=top align=center", "", "70%", "", "", "", BtnAva & "<div style=" & m_search_style2 & "><b><p><div id=infoPAV ><h5><h5></div></div>", "si")


    sTarget = ""






    Fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")





    celda = OFV.GenCelda("", "valign=top ", "", "", "", "", "", "temas", "si")




    celda = celda & OFV.GenCelda("", " valign=top ", "", "", "", "", "", "<b>" & UCase(SESQUEMAS) & "</b>", "si")

    '  Fila = fila &  ofv.GenRow("", "", "", "", "", "", "", celda, "si")

    tablabavan = tablabavan & OFV.GenTabla("", "aaa   STYLE='border-style:solid;border-width:1;border-color:#ffcc00;' ", "", "", "", "0", "", "0", "0", "", "", Fila, "si")

    tablabavan = tablabavan & "<br>" & "</DIV>" & Chr(13)

    response.write shtml & tablab & tablabavan
    response.flush
    Formulario = ""    'ofv.Encabezados("info")




    Set findd = Server.CreateObject("CorsearchExXml.Xmlado")

    Set cnbsq = OFV.conectar(strconn8)

    Set findd.OFV = OFV
    Set findd.Session = Session
    Set findd.response = response
    Set findd.Server = Server
    Set findd.Cnb = cnbsq

   

    findd.search_phrase = EliminarAcentos(search_phrase & "")

    If tipobsq <> "" Then
        findd.orden_doc = " ,CASE tipo "

        findd.orden_doc = findd.orden_doc & orden_doc
        findd.orden_doc = findd.orden_doc & " Else "
        findd.orden_doc = findd.orden_doc & " 99 "
        findd.orden_doc = findd.orden_doc & " End "
        findd.orden_doc = findd.orden_doc & " AS ORDEN_DOC "
    Else
        findd.orden_doc = " ,CASE tipo "
        findd.orden_doc = findd.orden_doc & orden_docFull
        findd.orden_doc = findd.orden_doc & " Else "
        findd.orden_doc = findd.orden_doc & " 99 "
        findd.orden_doc = findd.orden_doc & " End "
        findd.orden_doc = findd.orden_doc & " AS ORDEN_DOC "
    End If

    findd.canttx = 300
    findd.cantreg = 1000

'response.write findd.orden_doc
    If findd.IsNoiseWordex(findd.search_phrase) = 0 Then

        xinfo = "<h5>Sin Resultados..........</h5>"
        shtml = "<script>"


        shtml = shtml & " var op1=document.getElementById('infoP')" & Chr(13)

        shtml = shtml & " var op2=document.getElementById('infoPAV');" & Chr(13)

        shtml = shtml & " op1.innerHTML='"& xinfo&"';"& chr(13)
        shtml = shtml & " op2.innerHTML='"& xinfo&"';"& chr(13)

        shtml = shtml & "var ops1=document.getElementById('PAGINAS');ops1.innerHTML='"&xinfo&"'; "

        shtml = shtml & "var ops2=document.getElementById('PAGINASAV');ops2.innerHTML='"&xinfo&"'; "




        shtml = shtml & " </script> "

    Else


        Call findd.findDoc


        Set rsc = findd.orsFx
        '   RESPONSE.WRITE filtrost
        If tipobsq <> "" Then
            rsc.SORT = "ORDEN_DOC,rank desc,codigo,ranksol desc"



            If filtroEsq <> "" Then
                rsc.Filter = 0

                rsc.Filter = filtroEsq

            End If



            If filtrost <> "" Then

                rsc.Filter = filtrost
            End If

            If sINDOC Then
                rsc.Filter = "tipo='ZXZZZZX'"
            End If

        Else
            'rsc.SORT = "ORDEN_DOC,codigo,rank desc ,ranksol desc"
            RSC.SORT="ORDEN_DOC,rank desc,codigo ,ranksol desc"
        End If

        xo = 0
        url = "#Foo"
        totaldoc = 1
        ipag = 1
        Formularioz = ""
        ' rsc.move 11
        ' rsc.movenext 11
        solapa10 = False
        solapa5 = False
        Do While Not rsc.EOF And ipag < 11

          iddoc = rsc(0)
          if existeDoc(iddoc,rsc("tipo"),tipodoc) then
            solapa10 = False
            solapa5 = False
            xo = xo + 1
            xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
            xacceso = Replace(xacceso, "_", "&")
            If Portada(rsc("codigo")) = "SI" Then
                xacceso = Replace(xacceso, "*", "=") & "&NIV1=0&NIV2=" & rsc("index1")
            Else
                xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
            End If


            sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
            clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)

            if session("APLMDCLI")="BGAL" THEN
             Formularioz=  Formularioz & NewLine(0) & NewLine(0) & "<a  class=xxx onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('"&armatip(rsc(0))&"',250,'#dddddd')"" onclick="&clickn&" href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>" & tipodoc & " : " &  replace(rsc("title")," ","&nbsp;") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            ELSE
            Formularioz=  Formularioz & NewLine(0) & NewLine(0) & "<a  class=xxx onMouseout=""hideddrivetip()""  onMouseover=""ddrivetip('"&armatip(rsc(0))&"',250,'#dddddd')"" onclick="&clickn&" href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>" &  replace(rsc("title")," ","&nbsp;") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            END IF

            texto = NewLine(0) & "<br><span style=""" & m_style_text & """>" & rsc(3) & "</span>" & "<br>"
            i = 1
            atit = titsolapa(rsc(0), rsc("tipo"))
            '   response.write atit(1)
            antsol = 0



           if rsc("solapa") = 5 THEN
            solapa5 = TRUE
           END IF
           IF   rsc("solapa") = 10 then

              solapa10 = TRUE
           END IF

            rsc.movenext
            Do While Not rsc.EOF

                If iddoc <> rsc(0) Then
                    iddoc = rsc(0)
                    Formularioz = Formularioz & NewLine(0) & "<br>" & texto & "<hr>"
                    totaldoc = totaldoc + 1
                    Exit Do
                End If

                 If rsc("solapa") = "5" Then

				                    If Not solapa5 Then
				                        'if rsc("solapa") <> antsol and not solapa10 then
				                        xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
				                        xacceso = Replace(xacceso, "_", "&")
				                        xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
				                        sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
				                        clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
				                        Formularioz = Formularioz & "<a  onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('" & GETSOLAPA(rsc("solapa"), rsc("index1")) & "',450,'#dddddd')"" class=xxx onclick=" & clickn & " href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>&nbsp;" & i & "&nbsp;</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				                        i = i + 1
				                        ' antsol=rsc("solapa")
				                        solapa5 = True
				                    End If

                elseIf rsc("solapa") = 10 Then
                    If Not solapa10 Then
                        'if rsc("solapa") <> antsol and not solapa10 then
                        xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
                        xacceso = Replace(xacceso, "_", "&")
                        xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
                        sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
                        clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
                        Formularioz = Formularioz & "<a  onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('" & GETSOLAPA(rsc("solapa"), rsc("index1")) & "',450,'#dddddd')"" class=xxx onclick=" & clickn & " href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>&nbsp;" & i & "&nbsp;</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        i = i + 1
                        ' antsol=rsc("solapa")
                        solapa10 = True
                    End If
                Else
                    xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
                    xacceso = Replace(xacceso, "_", "&")
                    xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
                    sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
                    clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
                    Formularioz = Formularioz & "<a  onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('" & GETSOLAPA(rsc("solapa"), rsc("index1")) & "',450,'#dddddd')"" class=xxx onclick=" & clickn & " href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>&nbsp;" & i & "&nbsp;</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                    i = i + 1
                End If
                'end if

                rsc.movenext
                ' 0 i=i+1
            Loop
            If totaldoc = 9 Then
                totaldoc = 1
                If ipag = 1 Then
                    'Formularios=Formularios & "<div id=pag"&ipag&" >" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)
                    response.write "<div sTYLE='float: left; z-index:1;overflow:auto;width:" & anchopag & ";border-style:solid;border-width:2;border-color:#ffffff;' id=pag"&ipag&" >" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)
                    Formularioz = ""
                    response.flush
                Else
                    'Formularios=Formularios & "<div id=pag"&ipag&"  style='display=none;'>" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)
                    response.write  "<div id=pag"&ipag&"  sTYLE='float: left; z-index:1;overflow:auto;width:" & anchopag & ";border-style:solid;border-width:2;border-color:#ffffff;display:none;'>" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)
                    Formularioz = ""
                    response.flush
                End If
                ipag = ipag + 1
                Formularioz = ""
            End If
           else
             rsc.movenext
          end if
        Loop



        If totaldoc >= 1 Then
            If ipag = 1 And xo > 0 Then
                '     Formularios=Formularios & "<div id=pag"&ipag&" >" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)
                response.write  "<div sTYLE='float: left; z-index:1;overflow:auto;width:" & anchopag & ";border-style:solid;border-width:2;border-color:#ffffff;' id=pag"&ipag&" ><br>" & chr(13)  & formularioz & NewLine(0) &  "<br>" & texto & "<hr>" & chr(13)  &"</div>" & chr(13)
                response.flush
            Else
                '    Formularios=Formularios & "<div id=pag"&ipag&"  style='display=none;'>" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)

                if ipag < 11 then
                response.write "<div id=pag"&ipag&"  sTYLE='float: left; z-index:1;overflow:auto;width:" & anchopag & " ;border-style:solid;border-width:2;border-color:#ffffff;display:none;'>" & chr(13)  & formularioz & NewLine(0) &  "<br>" & texto & "<hr>" & chr(13)  &"</div>" & chr(13)
                response.flush
                else
                ipag=ipag-1
                end if
            End If
        End If




        If xo > 0 Then
        if ipag = 1 then

        xinfo = "<h5>Pagina 1" & " de " & ipag  & " - Total Documentos encontrados : " & xo & "<h5>"
        else
        xinfo = "<h5>Pagina 1" & " de " & ipag  & " - Total Documentos encontrados : " & xo & "<h5>"
        end if
        Else
            xinfo = "<h5> Total Documentos encontrados : " & xo & "<h5>"
        End If

        shtml = "<script>"

        '  if tipobsq="" then
        shtml = shtml & " var op1=document.getElementById('infoP')" & Chr(13)
        '   else
        shtml = shtml & " var op2=document.getElementById('infoPAV');" & Chr(13)
        '   end if
        shtml = shtml & " op1.innerHTML='"& xinfo&"';"& chr(13)
        shtml = shtml & " op2.innerHTML='"& xinfo&"';"& chr(13)
        shtml = shtml & "topepag=" & ipag & ";" & Chr(13)
        shtml = shtml & "cantdoc=" & xo & ";" & Chr(13)
        If xo > 0 Then
            '  if tipobsq="" then
            shtml = shtml & "var ops1=document.getElementById('PAGINAS');ops1.innerHTML='PAGINAS :&nbsp;&nbsp;'; "
            '  ELSE
            shtml = shtml & "var ops2=document.getElementById('PAGINASAV');ops2.innerHTML='PAGINAS :&nbsp;&nbsp;'; "
            '  END IF


            shtml = shtml & " for(var id=1;id <= topepag ;id++)"
            shtml = shtml & "{ "
            shtml = shtml & " if (id==1 && topepag !=1){"
            shtml = shtml & " ops1.innerHTML=ops1.innerHTML + '<a  id=lnk' + id  + ' href=""#Foo""   onclick=""mostrarp(' + id +');"" style=""" & m_style_titlex & """ >&nbsp' + id + '&nbsp</a>&nbsp;&nbsp;&nbsp;';" & Chr(13)
            shtml = shtml & "}else{"
            shtml = shtml & " ops1.innerHTML=ops1.innerHTML + '<a  id=lnk' + id  + ' href=""#Foo""   onclick=""mostrarp(' + id +');"" style=""" & m_style_title & """ >&nbsp' + id + '&nbsp</a>&nbsp;&nbsp;&nbsp;';" & Chr(13)
            shtml = shtml & "}"
            shtml = shtml & " ops2.innerHTML=ops2.innerHTML + '<a  id=lnkav' + id  + ' href=""#Foo""   onclick=""mostrarpav(' + id +');"" style=""" & m_style_title & """ >&nbsp' + id + '&nbsp</a>&nbsp;&nbsp;&nbsp;';" & Chr(13)
            shtml = shtml & "}"
        Else

            '  if tipobsq="" then
            shtml = shtml & "var ops1=document.getElementById('PAGINAS');ops1.innerHTML='&nbsp;&nbsp;'; "
            '  ELSE
            shtml = shtml & "var ops2=document.getElementById('PAGINASAV');ops2.innerHTML='&nbsp;&nbsp;'; "
            '  END IF
        End If

        shtml = shtml & " </script> "
    End If

    response.write shtml


End If

Function armatip(idC)
    texto = ""

   ' on error resume next
    strsql = "select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden from CSRELPATHDESCRIex a "
    strsql = strsql & " inner join objetos b1 on(b1.codobjeto= a.cobj)"
    strsql = strsql & " inner join objprty b on(b.codobjeto= a.cobj)"
    strsql = strsql & " inner join estados c on(c.objid=b.estado) where a.cobj=0" & idC
    '  RESPONSE.WRITE STRSQL
    Set recordsetip = CN.Execute(strsql)
    'response.write recordsetip("autor")
    'xxx=getinfouser(recordsetip("autor"))
    if not recordsetip.eof then
    texto = "<B><U>TEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("TEMADESCRI") & "<BR><B><U>SUBTEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("STEMADESCRI")    '& "<BR>"
    'TEXTO = TEXTO & "<B><U>AUTOR</U>:&nbsp;&nbsp;</B>" & Replace(Session("SGU_NOMBRE"), " ", "&nbsp;") '& "&nbsp;<B>SUBTEMA : </b>"& recordsetip("STEMADESCRI")& "<BR>"
    ' TEXTO = TEXTO &  "&nbsp;<B>SUBTEMA : </b>"& recordsetip("STEMADESCRI")& "<BR>"
    texto = texto & "<br><B><U>ESTADO</U>&nbsp;&nbsp;:</B>" & recordsetip("ESTADOD") & "<br><U><b>VIGENCIA</U> : </b>" & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "")
    If recordsetip("VERSION") = "O" Then
        zVersion = "(O)"
    Else
       xx= recordsetip("NROVER")

       zVersion = "V (" &  xx & ")"
    End If

     if session("APLMDCLI") <> "BGAL" THEN
   texto = texto & "<BR><B><U>COD. IDENT.</U> :&nbsp;&nbsp;</B>" & recordsetip("CODIDEN")
    end if
    texto = texto & "<BR><B><U>VERSIÓN</U>  :&nbsp;&nbsp;</B>" & zVersion & "  "    '& idC
    armatip = "<b>" & texto & "</b>"
    else
    armatip=""
    end if
    Call OFV.cerrarconsulta(recordsetip)
End Function


Function Portada(idC)
    texto = ""

    if session("APLMDCLI")="BC" THEN
    strsql = "select codobjeto from objportada where dettipo='S' and codobjeto=0" & idC

    Set recordseport = CN.Execute(strsql)

    If Not recordseport.EOF Then
        Portada = "SI"
    Else
        Portada = "NO"
    End If
     Call OFV.cerrarconsulta(recordseport)
    ELSE
    Portada = "NO"
    END IF

End Function



Private Function NewLine(spaces)
    NewLine = vbCrLf & Space(spaces)
End Function


Public Function ReplaceRegExp(ByVal strString, strPattern, strReplace)
    Dim RE: Set RE = New RegExp

    With RE
        .Pattern = strPattern
        .Global = True
        .IgnoreCase = True
        ReplaceRegExp = .Replace(strString, strReplace)
    End With
End Function

Function titsolapa(sCodigo, stipo)

    strsql = "Select * From Objperfildoc Where codobjeto =" & sCodigo
    'response.write "sql " & strsql
    Set rsperfil = OFV.crearconsultaEx(strsql, CN, 1, volver)
    If rsperfil.EOF Then
        perfilsp = "N"
        '  response.write "no "

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''' VERIFICA EL PERFIL GENERAL DEL OBJETO ''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        strsql = "Select Objperfil.* From Objperfil, Tipobjeto where "
        strsql = strsql & "Objperfil.tipoperfil = Tipobjeto.perfil "
        strsql = strsql & "and Tipobjeto.tipobjeto = '" & stipo & "'"
        '     response.write "sql 1111" & strsql
        Set rsperfil = OFV.crearconsultaEx(strsql, CN, 1, volver)


    End If

    If Not rsperfil.EOF Then
        stit = "vacio"
        delim = "###"
        For t = 17 To 33
            stit = stit & delim & rsperfil(t)

        Next
        '  response.write stit
        titsolapa = Split(stit, "###")
    End If
    Call OFV.cerrarconsulta(rsperfil)
End Function
Function GETSOLAPA(id, evt)
    ID=CSTR(ID&"")

    Select Case id

    Case "1"
        TIT = "SOLAPA " & atit(1)    '"Solapa propiedades"
    Case "2"
        TIT = "SOLAPA " & atit(2)    '"Solapa consideraciones"
    Case "3"
        TIT = "SOLAPA " & atit(3)    '"Solapa sectores"
    Case "4"
        TIT = "SOLAPA " & atit(4)    ' "Solapa eventos"

    Case "5"
        TIT = "SOLAPA " & atit(5)    '"Solapa procedimientos"
        TIT = TIT & "<b>" & bsqevt(evt) & "</b>"
    Case "6"
        TIT = "SOLAPA " & atit(6)    '"Solapa descripcion"
    Case "7"
        TIT = "SOLAPA " & atit(7)    '"Solapa controles"
    Case "8"
        TIT = "SOLAPA " & atit(8)    '"Solapa esquemas"
    Case "9"
        TIT = "SOLAPA " & atit(9)    '"Solapa tablas"
    Case "10"
        TIT = "SOLAPA " & atit(10)    '"Solapa documentos"
    Case "11"
        TIT = "SOLAPA " & atit(11)    '"Solapa contenido"
    Case "12"
        TIT = "SOLAPA " & atit(12)    '"Solapa cursograma"
    Case "13"
        TIT = "SOLAPA " & atit(13)    '"Solapa contratos"
    Case "14"
        TIT = "SOLAPA " & atit(14)    '"Solapa faq"
    Case "15"
        TIT = "SOLAPA " & atit(15)    '"Solapa general"
    Case Else
        TIT = "Sin descripción"



    End Select
    GETSOLAPA = "<B>" & TIT & "</B>"
End Function


Function bsqevt(id)

    sql = "select * from CSCEventoSectorPerfil Where codproced ='" & id & "'"

    Set recordevt = CN.Execute(sql)

    If Not recordevt.EOF Then
        EVSC = recordevt(2)
        EVPS = recordevt(6)
        EVDS = recordevt(4)
        secd = recordevt(3)
        perd = recordevt(5)
        FREC = recordevt(8)
        stexto1 = "<br><br>EVENTO: " & EVSC & "." & EVPS & "&nbsp;-&nbsp;" & EVDS
        stexto1 = stexto1 & "<br>SECTOR: " & secd
        stexto1 = stexto1 & "<br>PERFIL: " & perd
        stexto1 = stexto1 & "<br>FRECUENCIA: " & FREC
        stexto1 = Replace(stexto1, " ", "&nbsp;")
        bsqevt = stexto1

    End If

    Call OFV.cerrarconsulta(recordevt)


End Function


Function EliminarAcentos(texto)

    If Len(texto) <> 0 Then
        texto = reemplaza(texto, "[ÀÁÄáàä]", "A")
        texto = reemplaza(texto, "[ÈÉËèéë]", "E")
        texto = reemplaza(texto, "[ÌÍÏìíï]", "I")
        texto = reemplaza(texto, "[ÒÓÖòóö]", "O")
        texto = reemplaza(texto, "[ÙÚÜùúü]", "U")
    End If
    EliminarAcentos = texto
End Function

function existeDoc(id,nCodobjeto,BYREF TIPODOC)
dim strsql
dim recordsetip
 existeDoc=false



' select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden ,D.SEGOWN
 'from CSRELPATHDESCRIex a
' inner join objetos b1 on(b1.codobjeto= a.cobj)
' inner join objprty b on(b.codobjeto= a.cobj)
' inner join estados c on(c.objid=b.estado)
' inner join objsdk d on(d.codobjeto=b.codobjeto )
' left join  OBJSDKDET e on(
' E.CODOBJETO= D.CODOBJETO
' )
' where


' a.cobj=02202 and b1.objnivel >= 3  And c.objid In (9,1,6,2,5,3)
' AND
'((E.grpusr = 'USR-CL') or (E.grpusr = 'GRP-CONS'))


strsql = "select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden,E.DESCRI AS TIPODOC from CSRELPATHDESCRIex a "
    strsql = strsql & " inner join objetos b1 on(b1.codobjeto= a.cobj)"
    strsql = strsql & " inner join objprty b on(b.codobjeto= a.cobj)"
     strsql = strsql & " inner join tipobjeto e on(e.tipobjeto= b.tipobjeto)"
    strsql = strsql & " inner join estados c on(c.objid=b.estado) where a.cobj=0" & id
      if  session("APLMDCLI") = "BGAL" THEN
	      strsql=strsql & validadoc()
    end if
   'RESPONSE.WRITE STRSQL
    Set recordsetip = CN.Execute(strsql)
    if not recordsetip.eof then
    if  session("APLMDCLI") <> "BGAL" THEN
	     TIPODOC=recordsetip("TIPODOC")
	    if valpermisosdoc(nCodobjeto,id) then
	       existeDoc=true
	    else
	       existeDoc=false
	    end if
	    else
	    existeDoc=true
	    end if
	    else
	       existeDoc=false

    end if

    call ofv.cerrarconsulta(recordsetip)
end function


function valpermisosdoc(sTipobjeto,nCodobjeto)

   valpermisosdoc=false
   Strsql = "Select * From OBJSDK Where tipobjeto='" & sTipobjeto & "' and codobjeto = "  & nCodobjeto
'response.write strsql
   set rsval1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

   if not rsval1.eof then

       Segct = "S"
       Segmo = "S"
       Segco = "S"
       Segce = "S"

       if rsval1(3) then Segct = "N"
	   if rsval1(4) then Segmo = "N"
	   if rsval1(5) then Segco = "N"
       if rsval1(6) then Segce = "N"
       ownerDoc=rsval1(7)

       call ofv.cerrarconsulta(rsval1)

      if Session("Usuario") <> ownerDoc and Segco = "N" then


         Strsql = "Select * From OBJSDKDET Where tipobjeto='" & sTipobjeto
         Strsql = Strsql & "' and codobjeto = "  & nCodobjeto
         Strsql = Strsql & " and ("
         Strsql = Strsql & "(grpusr = 'USR-" & Session("Usuario") & "') "

'''''''''''''''''''''''''''''''''''''''''''
            if  Session("DELLOG") = "S" then

                XMFTGRP = split(session("SegGrpLst"),chr(9))
                xcount  = ubound(XMFTGRP) + 1
                XMFSEP  = ""
                StrSql = StrSql & " or ( grpusr in ("

                for XMFI = 0 to ubound(XMFTGRP)
                    StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
                    XMFSEP  = ","
                next

                StrSql = StrSql & ")) "

            else
                Strsql = Strsql & " or (grpusr = 'GRP-" & Session("Seggrp") & "')"
            end if


''''''''''''''''''''''''''''''''''''''''''


         Strsql = Strsql & ") "

        ' response.write strsql
         set rsval = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if  not rsval.eof then ' ok si esta en la lista  de consulta del doc
           valpermisosdoc=true
         else

             valpermisosdoc=false


         end if

          call ofv.cerrarconsulta(rsval)

    else ' si es el owner del doc

     valpermisosdoc=true

    end if

end if


end function

Function validadoc()


    sFormid = "CLIMS157172171124281051096140"

    ReDim Preserve bconsultar(4)
    OFV.ObtenerAtributosNodo sFormid, "", "ESTADO", "4"
    bconsultar(4) = OFV.AConsultar
    OFV.ObtenerAtributosNodo sFormid, "", "ESTADO", "2"
    bconsultar(2) = OFV.AConsultar
    OFV.ObtenerAtributosNodo sFormid, "", "ESTADO", "1"
    bconsultar(1) = OFV.AConsultar
    OFV.ObtenerAtributosNodo sFormid, "", "ESTADO", "3"
    bconsultar(3) = OFV.AConsultar
    sverestado = " and b1.objnivel >= " & Session("SegNivel") & " And c.objid In (9,"


    For i = 1 To 4
        If bconsultar(i) Then
            sverestado = sverestado & i & ","
            If i = 1 Then sverestado = sverestado & "6" & ","
            If i = 2 Then sverestado = sverestado & "5" & ","

        End If
    Next
    sverestado = Left(sverestado, Len(sverestado) - 1) & ") "


    validadoc = sverestado
End Function


Function reemplaza(texto, patrn, replStr)
    Dim regEx
    Set regEx = New RegExp            ' Crea una expresión regular.
    regEx.Pattern = patrn            ' Establece el modelo.
    regEx.IgnoreCase = True            ' Hace distinción entre mayúsculas y minúsculas.
    regEx.Global = True
    reemplaza = regEx.Replace(texto, replStr)   ' Realiza el reemplazo.
    Set regEx = Nothing
End Function
%>
