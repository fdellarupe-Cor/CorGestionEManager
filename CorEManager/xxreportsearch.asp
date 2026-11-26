
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTESTRUCTURA.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTGENERAL.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTPROCEDIMIENTOS.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTCONTROLES.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTESQUEMAS.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTTABLAS.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTARCHIVOS.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTSECTORES.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTWEDT.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTOBJDSC.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTFAQ.inc" -->

<%

stipo = Request.querystring("tipobj")
sCodigo = Request.querystring("id")
SSOLAPA = Request.querystring("SOLAPA")
sindice = Request.querystring("INDICE")
aplcheck()
ncodigo = sCodigo


Server.ScriptTimeout = 60 * 7

Session("usuario") = "GOOGLE"

ok = OFV.CheckUsuario()
If ok <> "Y" Then
    uv = OFV.uv
    If uv = "NO" Then ok = OFV.ValidarUsuario
    shtml = ok
Else

    response.expires = 0




    Session("ModMain") = "ED"


    xapli = 1





    ''''''''''''''''''''''''''''



    shtml = "<html><head>"








    Set CN = OFV.conectar(OFV.strconn2)

    ''''''''''''''''''''''''''''''''''''''''''''

    WKPROCDESC = ""

    strsql = "Select A.*,b.descri as tobjdsc From OBJETOS A"
    strsql = strsql & " INNER JOIN tipobjeto B ON ( B.tipobjeto = A.tipobjeto )"
    strsql = strsql & " where  a.codobjeto = '" & sCodigo & "' "


    DESCRIABR = ""
    Set rs = OFV.crearconsultaEx(strsql, CN, 1, parametros)
    If Not rs.EOF Then

        WKPROCDESC = rs("tobjdsc") & ":&nbsp;" & rs(3)
        WKPROCDESC2 = WKPROCDESC
        DESCRIABR = rs("descriabr")
    End If

    Call OFV.cerrarconsulta(rs)

    ' WKPROCDESC="&nbsp;"

'response.write WKPROCDESC
    celda = ""
    celda = "<title>" & WKPROCDESC & "</title><meta name='robots' content='noarchive'></head><body>"

    TablaH = celda



    Fila = ""

    wspcolor1 = "#990000"
    wspcolor2 = "#009900"
    sImageO1 = "&#0253;"
    sImageO2 = "&#0254;"
    xTrColor = "#FFFFFF"


    strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & ncodigo
    'response.write strsql
    Set rs = OFV.crearconsultaEx(strsql, CN, 1, parametros)
    bVersion = "NO"
    If Not rs.EOF Then
        If rs(0) = "V" Then
            bVersion = "SI"
        End If
        Nestado = rs(1)
    End If
    Call OFV.cerrarconsulta(rs)

    zztabla = ""
    delimsol = ""
    XDATA = "|REG|"
    delimsolC = "|SOLAPA|"
    delimsol = delimsolC

    bSector = False
    bEvento = False
    bProcedimiento = False
    bControl = False
    bEsquema = False
    bTablas = False
    bDocum = False
    bEditor = False
    bDescri = False
    bMain = False
    bcontratos = False
    bpreguntas = False
    bgeneral = False


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
        Set rsperfil = OFV.crearconsultaEx(strsql, CN, 1, volver)

    Else
        ' response.write "si "
        perfilsp = "s"

    End If

    If Not rsperfil.EOF Then

        If rsperfil("objsector") Then bSector = True
        'Sectores

        If rsperfil("objeven") Then bEvento = True
        'Eventos

        If rsperfil("objproced") Then bProcedimiento = True
        'Procedimientos

        If rsperfil("objcontrol") Then bControl = True
        'Controles

        If rsperfil("objesquema") Then bEsquema = True
        'Esquemas

        If rsperfil("objtabla") Then bTablas = True
        'Tablas

        If rsperfil("objhtm") Then bDocum = True
        'Documentos

        If rsperfil("objobj1") Then bEditor = True

        'CONTRATOS

        If rsperfil("objobj3") Then bcontratos = True

        'PREGUNTAS FRECUENTES
        If rsperfil("objobj4") Then bpreguntas = False


        'Editor estructurado

        If rsperfil("objaccion") Then bDescri = True
        'Descripcion

        If rsperfil("objmain") Then bMain = True
        'Consideraciones Generales

        If rsperfil("objmain") Then bMain = True
        'Consideraciones Generales

        If rsperfil("objobj5") Then bgeneral = True
        'solapa General

    End If


    Call OFV.cerrarconsulta(rsperfil)

    xentro = False

    'for isolapas = 1 to 15
    'SSOLAPA= isolapas
    'nid=0
    'SELECT CASE SSOLAPA

    '    CASE 1 'propiedades

    '   CASE 2 'consideraciones
    tabla = ""
    If bMain Then
        id = ncodigo
        SSOLAPA = 2
        If bgeneral = False And xentro = False Then
            tabla = XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTWEDT(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
            xentro = True
        Else
            tabla = XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTWEDT(ncodigo, stipo) & delimsol
        End If
        delimsol = delimsolC
    End If
    '   CASE 3 'sectores

    If bSector Then
        SSOLAPA = 3
        If bgeneral = False And xentro = False Then
            tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTSECTORES(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
            xentro = True
        Else
            tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTSECTORES(ncodigo, stipo) & delimsol
        End If

        delimsol = delimsolC
    End If
    '   CASE 4 'eventos

    '   CASE 5 'PROCEDIMIENTOS
    If bProcedimiento Then
        'nID = sindice
        SSOLAPA = 5
        xsql = "select distinct(codproced)  from OBJACCION where  codobjeto=0" & ncodigo

        sCambios = "NO"
        Set rscx = OFV.crearconsultaEx(xsql, CN, 1, parametros)
        'Tabla=""
        Do While Not rscx.EOF
            nID = rscx(0)    'sindice
            If bgeneral = False And xentro = False Then
                tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#" & nID & XDATA & GENFMTPROCEDIMIENTOS(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
                xentro = True
            Else
                tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#" & nID & XDATA & GENFMTPROCEDIMIENTOS(ncodigo, stipo) & delimsol
            End If
            delimsol = delimsolC
            rscx.movenext
        Loop
        Call OFV.cerrarconsulta(rscx)
    End If
    '   CASE 6 'descripcion

    If bDescri Then
        SSOLAPA = 6
        If bgeneral = False And xentro = False Then
            tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTOBJDSC(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
            xentro = True
        Else
            tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTOBJDSC(ncodigo, stipo) & delimsol
        End If
        delimsol = delimsolC
    End If

    '  CASE 7 'controles

    ' nID = sindice
    If bControl Then
        sCambios = "NO"
        SSOLAPA = 7
        xsql = "select objid  from objcontrol where  codobjeto=0" & ncodigo

        sCambios = "NO"

        Set rscx = OFV.crearconsultaEx(xsql, CN, 1, parametros)
        ' Tabla=""
        Do While Not rscx.EOF
            nID = rscx(0)    'sindice
            If bgeneral = False And xentro = False Then
                tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#" & nID & XDATA & GENFMTCONTROLES(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
                xentro = True
            Else
                tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#" & nID & XDATA & GENFMTCONTROLES(ncodigo, stipo) & delimsol
            End If
            delimsol = delimsolC
            rscx.movenext
        Loop
        Call OFV.cerrarconsulta(rscx)
    End If
    '  CASE 8 'esquemas
    If bEsquema Then
        SSOLAPA = 8
        nID = sindice

        sCambios = "NO"
        xsql = "select objid  from OBJESQUEMA2 where  codobjeto=0" & ncodigo
        sCambios = "NO"
        Set rscx = OFV.crearconsultaEx(xsql, CN, 1, parametros)
        ' Tabla=""
        Do While Not rscx.EOF
            nID = rscx(0)    'sindice
            If bgeneral = False And xentro = False Then
                tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#" & nID & XDATA & GENFMTESQUEMAS(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
                xentro = True
            Else
                tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#" & nID & XDATA & GENFMTESQUEMAS(ncodigo, stipo) & delimsol
            End If
            delimsol = delimsolC
            rscx.movenext
        Loop
        Call OFV.cerrarconsulta(rscx)
    End If
    '  CASE 9 'tablas
    If bTablas Then
        SSOLAPA = 9
        'Bloque de Defaults
        ncolcant = 1
        nBorde = False
        nblkfc = "#993333"
        nblkbc = "#ffffff"
        nblkff = "ARIAL"
        nblkfs = "10pt"

        nctitfc = "#000077"
        nctitbc = "#ffffff"
        nctitff = "ARIAL"
        nctitfs = "8pt"

        ncconfc = "#336699"
        ncconbc = "#ffffff"
        ncconff = "ARIAL"
        ncconfs = "8pt"
        ReDim aColumnas(0)

        nStilo = " Style={font-family:" & nblkff & ";font-size:" & nblkfs & ";color:#993333;background-color:#ffffff;font-weight:bold;} "

        nfontx = "</font>"

        xsql = "select objid  from Objtabdef where  codobjeto=0" & ncodigo
        sCambios = "NO"
        Set rscx = OFV.crearconsultaEx(xsql, CN, 1, parametros)
        ' Tabla=""
        Do While Not rscx.EOF
            nTabla = rscx(0)    'sindice

            strsql = "SELECT * From Objtabdef Where objid = " & nTabla
            '  response.write Strsql
            Set rst = OFV.crearconsultaEx(strsql, CN, 1, parametros)

            If Not rst.EOF Then
                nTitulo = rst(4)
                If nTitulo = "" Then nTitulo = "&nbsp;"
                nPerfil = rst(3)
            End If

            If bgeneral = False And xentro = False Then
                tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#" & nTabla & XDATA & GENFMTTABLAS(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
                xentro = True
            Else
                tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#" & nTabla & XDATA & GENFMTTABLAS(ncodigo, stipo) & delimsol
            End If
            delimsol = delimsolC
            rscx.movenext
        Loop
        Call OFV.cerrarconsulta(rscx)
    End If
    '  CASE 10 'documentos
    If bDocum Then
        SSOLAPA = 10
        SKIND = 1

        If bgeneral = False And xentro = False Then
            tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTARCHIVOS(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
            xentro = True
        Else
            tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTARCHIVOS(ncodigo, stipo) & delimsol
        End If
        delimsol = delimsolC
        ' Tabla = GENFMTARCHIVOS2(nCodigo,STIPO)
    End If
    '  CASE 11 'contenido

    If bEditor Then
        SSOLAPA = 11
        Marca = "SI"
        id = 0

        '''''''''''''''''''''''''''



        Dim Pdef(13, 18)

        strsql = "Select * From SysEstprty"
        Set rs = OFV.crearconsultaEx(strsql, CN, 1, parametros)
        t = 0
        Do Until rs.EOF
            For ti = 1 To 18
                If (rs(ti) & "") = "" Then
                    sti = ""
                Else
                    sti = rs(ti)
                End If

                Pdef(t, ti - 1) = sti

            Next
            rs.movenext
            t = t + 1
        Loop

        Call OFV.cerrarconsulta(rs)


        xsql = "select objid  from Objestr where  codobjeto=0" & ncodigo
        sCambios = "NO"
        Set rscx = OFV.crearconsultaEx(xsql, CN, 1, parametros)
        '   Tabla=""
        Do While Not rscx.EOF
            id = rscx(0)    'sindice

            strsql = "Select * From Objestr Where objid = " & id
            ''' response.write strsql & "<br>"
            Set rs = OFV.crearconsultaEx(strsql, CN, 1, parametros)
            bsecuen = ""
            bDescri = " "
            bobj = ""
            If Not rs.EOF Then
                bsecuen = rs(4)
                bDescri = rs(3)
                bobj = rs(0)

                stexto1 = "<A ID='ATIT" & rs(0) & "' NAME='ATIT" & rs(0) & "' " & Chr(13)
                stexto1 = stexto1 & " class=nt href='coestrCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
                stexto1 = stexto1 & "&sdocum=" & sdocum & "&sdocc=" & rs(0)
                stexto1 = stexto1 & "'   >" & sfidkc & "&nbsp;<big>" & rs(3) & "</big></a>" & Chr(13)

                Celdat = "<td class=ut style='cursor:hand;'  onclick='document.all.ATIT" & rs(0) & ".click();' >" & stexto1 & "</td>"
                filat = filat & OFV.GenRow("", "", "", "", "", "", "", Celdat, "si") & Chr(13)

            End If

            Call OFV.cerrarconsulta(rs)



            strsql = "Select * From ObjEstPrty where detcon = " & id
            Set rs = OFV.crearconsultaEx(strsql, CN, 1, parametros)
            t = 0
            Do Until rs.EOF
                Select Case rs(4)
                Case "A"
                    t = 0
                Case "B"
                    t = 1
                Case "C"
                    t = 2
                Case "D"
                    t = 3
                Case "E"
                    t = 4
                Case "F"
                    t = 5
                Case "G"
                    t = 6
                End Select

                For ti = 5 To 22
                    If (rs(ti) & "") = "" Then
                        sti = ""
                    Else
                        sti = rs(ti)
                        Pdef(t, ti - 5) = sti
                    End If



                Next
                rs.movenext
            Loop

            Call OFV.cerrarconsulta(rs)

            PInden = Pdef(0, 7)
            If PInden <> "NO" Then PInden = "SI"

            PNumer = Pdef(0, 12)
            If PNumer <> "NO" Then PNumer = "SI"

            PNumCls = Pdef(0, 14)
            ReDim Preserve PNumTip(13)
            For i = 0 To 13
                PNumTip(i) = Pdef(i, 13)
                If PNumTip(i) = "" Then PNumTip(i) = "1"
            Next

            '''''''''''''''''''''''''''
            If bgeneral = False And xentro = False Then
                tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#" & id & XDATA & GENFMTESTRUCTURA(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
                xentro = True
            Else
                tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#" & id & XDATA & GENFMTESTRUCTURA(ncodigo, stipo) & delimsol
            End If
            delimsol = delimsolC



            rscx.movenext
        Loop
        Call OFV.cerrarconsulta(rscx)

    End If
    '   CASE 12 'cursograma

    '   CASE 13 'contratos


    If bcontratos Then
        SKIND = 2
        SSOLAPA = 13
        If bgeneral = False And xentro = False Then
            tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTARCHIVOS(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
            xentro = True
        Else
            tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTARCHIVOS(ncodigo, stipo) & delimsol
        End If
        delimsol = delimsolC
    End If
    '    CASE 14 'faq
    If bpreguntas Then
        SSOLAPA = 14
        If bgeneral = False And xentro = False Then
            tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTFAQ(ncodigo, stipo) & "  " & WKPROCDESC & delimsol
            xentro = True
        Else
            tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTFAQ(ncodigo, stipo) & delimsol
        End If
        delimsol = delimsolC
    End If
    '    CASE 15 'general

    If bgeneral Then
        SSOLAPA = 15
        tabla = tabla & XDATA & ncodigo & "#" & stipo & "#" & SSOLAPA & "#0" & XDATA & GENFMTGENERAL(ncodigo, stipo) & DESCRIABR & "  " & WKPROCDESC & delimsol
        delimsol = delimsolC
    End If




    Call OFV.cerrarconn(CN)

    '''''''''''''''''''''''''''




    shtml = shtml & tabla


    ' sHTML=sHTML & "<script> document.body.innerHTML=document.body.innerText;</script>"
    shtml = shtml & "</body></html>"

End If

response.write shtml
''------------------

Function GENFMTARCHIVOS2(FNCODOBJ, FNTIPOBJ)


    FNSHTML = ""


    StrSqlx = "Select Version,estado,tipobjeto From Objprty Where Objprty.codobjeto = " & ncodigo
    Set rs = OFV.crearconsultaEx(StrSqlx, CN, 1, parametros)
    bVersion = "NO"
    If Not rs.EOF Then
        If rs(0) = "V" Then bVersion = "SI"
        Nestado = rs(1)

    End If
    Call OFV.cerrarconsulta(rs)


    sarch = "NO"

    strsql = "SELECT descri, archivo, objid, directorio "
    strsql = strsql & " from ObjHtm Where codobjeto = " & ncodigo    '& " AND TIPO ='" & SKIND &"'         "
    strsql = strsql & " order by secuen "

    Formulario = ""
    ' response.write strsql
    Set rs = OFV.crearconsultaEx(strsql, CN, 1, parametros)
    If Not rs.EOF Then
        Do Until rs.EOF

            sdirec = rs(3) & ""

            If sdirec = "" Then sdirec = "Archivos"




            sarch = sdirec & "/" & rs(1)
            clase = "<font color=#993333><u>"
            stexto = sfida & sfta & Replace(rs(0), " ", "&nbsp;") & sftc

            stexto1 = "<A class=ttt  href='" & sarch&"'"

            stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"

            If (rs(1) & "") = "" Then
                stexto1 = "<A class=ttt  href='javascript:void(0);' "
                stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"
            End If

            celda = ""


            ' estado observacion

            Celdaobs = OFV.GenCelda("", "ut", "", "1%", "", "", "", "&nbsp;", "si")




            If Nestado = 5 Then


                sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta & "&Codobjeto=" & ncodigo
                sAccion = sAccion & "&Tipobjeto=" & stipo & "&ID=" & rs(2) & "&ID2=" & rs(2)

                xsffun = " onclick=""document.getElementById('DX4').style.display='block';"" "
                sHref = sAccion
                sInput = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
                Celdaobs = OFV.GenCelda("", "ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ", "", "", "1", "", "", sInput, "si")
                Celdaobs = OFV.GenLink("", "ut" & "  ", sHref, "WK4", "", "", Celdaobs, "si")



            End If


            celda = celda & Celdaobs


            celda = celda & OFV.GenCelda("", "ut  ", "", "95%", "", "", "", stexto1, "si")
            Formulario = celda
            Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Formulario, "si")

            rs.movenext
        Loop

        FNSHTML = FNSHTML & "<br><br><center>" & OFV.GenTabla("", "aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-COLOR:#dddddd;' ", "", "90%", "", "0", "0", "4", "0", "", "", Fila, "si")
        '' if not sdocum = "" then
        If sarch <> "NO" Then
            sAccion = "CoArchivos3.asp?Archivo=" & sarch
            FNSHTML = FNSHTML & OFV.GenForm("cons", "", sAccion, "_blank", "post", "", sInput, "si")
        End If

    Else

        sdocus = "NO"


    End If

    '''SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where codcomp

    Fila = ""

    sarch = 0
    'strsql = "SELECT DISTINCT COMPLEMENTOS.codcomp, COMPLEMENTOS.codiden, "
    'strsql = strsql & "COMPLEMENTOS.descri, CSACOMPLE.nombre FROM COMPLEMENTOS, "
    'strsql = strsql & " OBJACCION, CSACOMPLE where COMPLEMENTOS.codcomp = OBJACCION.codcomp "
    ' strsql = strsql & " and COMPLEMENTOS.codcomp = CSACOMPLE.codcomp "
    'strsql = strsql & " and OBJACCION.codobjeto =" & ncodigo
    'strsql = strsql & " ORDER BY COMPLEMENTOS.codcomp "


    strsql = "Select distinct a.codobjetol, b.codiden, b.nombre , b.nombre From objlnk a, csacomple b  Where a.ltipo = 'KL' and a.codobjeto = 0" & ncodigo
    strsql = strsql & " and a.codobjetol = b.codcomp "
    strsql = strsql & " ORDER BY a.codobjetol"
    'response.write strsql
    Set rs = OFV.crearconsultaEx(strsql, CN, 1, "")
    Formulario = ""
    Set rs = OFV.crearconsultaEx(strsql, CN, 1, parametros)
    If Not rs.EOF Then
        Do Until rs.EOF


            If sdocum = CStr(rs(0)) Then
                sarch = rs(0)
                clase = "<font color=#993333><u>"
                stexto = sfida & sfta & Replace(rs(3) & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(2), " ", "&nbsp;") & sftc
            Else
                clase = ""
                stexto = sfidc & sft & Replace(rs(3) & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(2), " ", "&nbsp;")
            End If
            stexto1 = "<A class=ttt  href=CoArchCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
            stexto1 = stexto1 & "&sdocum=" & rs(0)
            stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"


            celda = OFV.GenCelda("", "ut  ", "", "1%", "", "", "", stexto1, "si")
            Formulario = celda

            rs.movenext

            Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Formulario, "si")
        Loop

        FNSHTML = FNSHTML & "<br><br><center>" & OFV.GenTabla("", "aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-COLOR:#dddddd;' ", "", "90%", "", "0", "0", "4", "0", "", "", Fila, "si")

        If sarch > 0 Then
            sAccion = "EditarImagen2.asp?ID=" & sarch
            FNSHTML = FNSHTML & OFV.GenForm("cons", "", sAccion, "_blank", "post", "", sInput, "si")
        End If

    Else
        scomple = "NO"
    End If
    Call OFV.cerrarconsulta(rs)


    If sdocum <> "" Then
        FNSHTML = FNSHTML & "<script language=javascript >"
        FNSHTML = FNSHTML & " document.cons.submit(); "
        FNSHTML = FNSHTML & "</script>"
    End If




    If scomple = "NO" And sdocus = "NO" Then
        sestr = "<h5>Sin&nbsp;datos&nbsp;disponibles"
        sestr = sestr & "</h5>"
        Formulario = OFV.GenCelda("", "ttt", "center", "80%", "", "", "", sestr, "si")
        Fila = OFV.GenRow("", "", "", "", "", "", "", Formulario, "si")
        FNSHTML = FNSHTML & "<br><br><center>" & OFV.GenTabla("", "aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-COLOR:#dddddd;' ", "", "90%", "", "0", "0", "4", "0", "", "", Fila, "si")
    End If


    GENFMTARCHIVOS2 = FNSHTML


End Function



%>

