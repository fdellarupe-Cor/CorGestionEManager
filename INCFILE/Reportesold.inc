<%



Class ClsReportes

Private m_oOFV 
Private m_oServerVar 
Private m_oRequestVar 
Private m_oResponseVar 
Private m_oSessionVar 


Private MAXHEIGHT, m_nAltoPagina
'Private oForm As Form
'Private EventType As LogEventTypeConstants

Public sHTML, sHead, sBloque1, sError, sRepHead, sPath2
Public nPag
Public RCliLogo 
Public RCliNomb 




Public Function UserV()

    UserV = OFV.uv

End Function

Public Property Set OFV(ByVal oOFV )

    Set m_oOFV = oOFV

End Property

Public Property Get OFV() 

    Set OFV = m_oOFV

End Property



Public Function Generar() 

   

    Dim id, tipo, nid, stipo, STitulo, sPerfil, sNrover, sPath, nEstado
    Dim sEsquema, sEsquemaEx, nDay, nMonth, nYear, EID, bsecuen
    Dim bDescri, PInden, Pnumer, Pnumcls, sclave, stip, smodi, sCompCopia
    Dim stipmodi, pfont, pfsize, pfcol, pbgcol, PBorde, PFtipo, lConten
    Dim PSang, PEspa, Palign, Ptalig, pStyle, pStyleF, BnumS, sreduc, sbord, salng
    Dim PVFont, PVtipo, pVind, sEvento, sSector, sprocsec, xsec, mpvin
    Dim sVerbo, sAccion, xSecuencia, sReferencia, sComplemento
    Dim sControlRef, sControlcon, sFrecuencia, nTabla, nTitulo
    Dim nPerfil, nblkff, nblkfs, nSec, nBTitulo, sConten
    Dim strsql, bVersion, sTema, sSubtema, sEstado, sFecha, sImagen
    Dim sftvi, sFechaVi, sftel, sFechaEl, sftem, sFechaEm, wlink
    Dim sftve, sFechaVe, hoja, topeflush, Celda, xVer, xWlnkPtr
    Dim bi, sObjetivo, sAlcance, sConsideracion, rtfFile, sh, xalig

    Dim TConsideracion, TObjetivo, TAlcance

    Dim xxx, aaa, bbb, Fila, t, tx, txi, ta, ti, sti, i
    Dim sHTML, tabla1, sant, bseci, bsecix, sForm, sImage
    Dim BsecD, bspan, pborst, PBCol, PFtipoS, PFtipoF, vt, y
    Dim pfonst, PnumS, PsangS, Vs, PVin, PVindS, Vi
    Dim ssale, oc, ocd, z, smark, bDupli, sCompAn, sControl
    Dim bPrimero, sdesc, Nblanco, nStilo, ncolcant, nblkfc
    Dim nBorde , nblkbc, nctitfc, nctitbc, nctitff
    Dim nctitfs, ncconfc, ncconbc, ncconff, ncconfs, ix
    Dim ltitu, nStilot, nStiloc, nFontr, ssize
    Dim nfontx, nfonty, nfontt, nfontc, nfonta
    Dim wrtexto, wrtextosplit, tenc, sInput
    Dim irtextoc, irtextoF, irtextoO , irtextoI
    Dim cn , rs , fso , tf 
    Dim Parser , rsx , rsa 
    Dim rsb , rsc , rsd , rse 
    Dim rsf , rsCuenta , rs1 
    Dim rs2 , rsalt 
    Dim rsz  
    Dim cnn 
    Dim wtipoLc
    Dim codobjetoL
    Dim sActcon
    Dim sHref
    Dim wNomL
    Dim sFuncion, sLink, XDescri, sTextoV
    Dim Wlinkd
    Dim rst , rst2 
    Dim StrArray 
    Dim IArray 
    Dim orsx 
    Dim ssql 
    Dim ResStr, ReferenciaLoop 
        

Generar = ""

Response.buffer = True
Server.scripttimeout = 600
id = Request.querystring("ID")
tipo = Request.querystring("tipobj")
STitulo = UCase(OFV.ConvertirCar(Request.querystring("Titulo"), "_", " "))
sPerfil = Request.querystring("Perfil")

Set cn = OFV.conectar(OFV.strconn2)

strsql = "Select Version,Nrover From Objprty Where Objprty.codobjeto = " & id
Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
bVersion = "NO"
If Not rs.EOF Then
    If rs(0) = "V" Then
        bVersion = "SI"
        sNrover = rs(1)
    End If
End If
Call OFV.cerrarconsulta(rs)

'840 sPath = Server.mappath(Request.ServerVariables("SCRIPT_NAME"))
'850 sPath2 = Mid(sPath, 1, InStrRev(sPath, "\", -1))
'860 sPath = sPath2

nid = id
stipo = tipo

strsql = "Select tipobjeto From objetos Where codobjeto = " & nid
Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
If Not rs.EOF Then
    stipo = rs(0)
End If
Call OFV.cerrarconsulta(rs)



Do
    strsql = "Select codobjmaster,tipobjmaster From Relaobj Where codojasoc = " & nid & " and tipobjasoc = '" & stipo & "'"
    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
    If Not rs.EOF Then
        nid = rs(0)
        stipo = rs(1)
    Else
        Exit Do
    End If
   
    Call OFV.cerrarconsulta(rs)
Loop
 'response.write strsql
strsql = "Select Tema.descri,Prodserv.descri From Tema, Prodserv "
strsql = strsql & "where Tema.codtema = Prodserv.codtema "
strsql = strsql & " and Prodserv.codprodserv = " & nid
Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
'response.write strsql
If Not rs.EOF Then
    sTema = rs(0)
    sSubtema = rs(1)
End If
Call OFV.cerrarconsulta(rs)

Dim sAutor
Dim sHead1
Dim sAutorDes
Dim xind
Dim nestadoD
   '    Dim sEstado
   '    Dim sFecha
   '    Dim sftvi
     
 RCliLogo = Session("clilogop")
 RCliNomb = Session("cliente")

strsql = "SELECT OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.descri, OBJPRTY.estado, "
strsql = strsql & "OBJPRTY.fecvigen, OBJPRTY.fecbaja, OBJPRTY.fecemis, OBJPRTY.fecvto, "
strsql = strsql & "OBJPRTY.autor, estados.estado, objetos.objUN "
strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, estados "
strsql = strsql & "where OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
strsql = strsql & " OBJETOS.codobjeto = OBJPRTY.codobjeto and "
strsql = strsql & " OBJPRTY.estado = estados.objid "
strsql = strsql & " and OBJETOS.codobjeto = " & id
'response.write strsql
Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
If Not rs.EOF Then
    sAutor = UCase(rs(8))
    sEsquema = rs(2)
    sEsquemaEx = rs(0) & " - " & rs(1)
    nEstado = rs(3)
    nestadoD = rs(9)
    sEstado = "&nbsp;"
    sFecha = "&nbsp;"
    sftvi = "Vigencia:"

     If rs(10) > 1 Then
        strsql = "Select * from clienteplan where id = " & rs(10)
        Set rsc = OFV.crearconsultaEx(strsql, cn, 1, "")
        If Not rsc.EOF Then
            RCliLogo = "../COREss20/Empresa/" & rsc(5)
            if (rs(11) & "" ) <> "" then RCliLogo = "../COREss20/Empresa/" & rsc(11)
            RCliNomb = rsc(1)
        End If
        Call OFV.cerrarconsulta(rsc)

     End If

    If CLng(rs(4)) = 0 Then
        sFechaVi = "&nbsp;"
    Else
        sFechaVi = CDate(rs(4) + CLng(CDate("01-01-1900")))
    End If



    sftel = "Eliminacion:"

    If CLng(rs(5)) = 0 Then
        sFechaEl = "&nbsp;"
    Else
        sFechaEl = CDate(rs(5) + CLng(CDate("01-01-1900")))
    End If

    sftem = "Emision:"

    If CLng(rs(6)) = 0 Then
        sFechaEm = "&nbsp;"
    Else
        sFechaEm = CDate(rs(6) + CLng(CDate("01-01-1900")))
    End If


    sftve = "Vencimiento:"


    If CLng(rs(7)) = 0 Then
        sFechaVe = "&nbsp;"
    Else
        sFechaVe = CDate(rs(7) + CLng(CDate("01-01-1900")))
    End If
End If
Call OFV.cerrarconsulta(rs)

sHead1 = OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")

Celda = OFV.GenCelda("", "iett", "", "10%", "", "", "", "Tema: ", "si")
Celda = Celda & OFV.GenCelda("", "ietx", "", "70%", "", "", "", sTema, "si")
Celda = Celda & OFV.GenCelda("", "iett", "", "10%", "", "", "", sftem, "si")
Celda = Celda & OFV.GenCelda("", "ietx", "", "10%", "", "", "", fechadorCor(sFechaEm,""), "si")
sHead1 = sHead1 & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

Celda = OFV.GenCelda("", "iett", "", "10%", "", "", "", "Subtema: ", "si")
Celda = Celda & OFV.GenCelda("", "ietx", "", "70%", "", "", "", sSubtema, "si")
Celda = Celda & OFV.GenCelda("", "iett", "", "10%", "", "", "", sftvi, "si")
Celda = Celda & OFV.GenCelda("", "ietx", "", "10%", "", "", "", FechadorCor(sFechaVi,""), "si")
sHead1 = sHead1 & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

If bVersion = "SI" Then
    xVer = "Version: " & sNrover
Else
    xVer = "Original"
End If

Celda = OFV.GenCelda("", "iett", "", "80% colspan=2", "", "", "", xVer, "si")
Celda = Celda & OFV.GenCelda("", "iett", "", "10%", "", "", "", sftve, "si")
Celda = Celda & OFV.GenCelda("", "ietx", "", "10%", "", "", "", fechadorCor(sFechaVe,""), "si")
sHead1 = sHead1 & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

Celda = OFV.GenCelda("", "iett", "", "10%", "", "", "", "Estado: ", "si")
Celda = Celda & OFV.GenCelda("", "ietx", "", "70%", "", "", "", nestadoD, "si")
Celda = Celda & OFV.GenCelda("", "iett", "", "10%", "", "", "", sftel, "si")
Celda = Celda & OFV.GenCelda("", "ietx", "", "10%", "", "", "", fechadorCor(sFechaEl,""), "si")
sHead1 = sHead1 & OFV.GenRow("", "", "", "", "", "", "", Celda, "si") & OFV.GenfTabla()

sHead = OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")

Celda = OFV.GenCelda("", "iett", "right", "100%", "", "", "", "Pagina:&nbsp;&&XX de &&XY", "si")
sHead = sHead & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
sHead = sHead & OFV.GenRow("", "", "", "", "", "", "", Celda, "si") & OFV.GenfTabla()


sRepHead = RepHeader2(sEsquema & " " & sEsquemaEx, "600")
sHTML = ""

ReDim Bloque(8)
For bi = 0 To 8
    Bloque(bi) = False
Next

ReDim TBloque(8)
For bi = 0 To 8
    TBloque(bi) = " "
Next

strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
If Not rs.EOF Then
    If rs(3) Then Bloque(1) = True
    If rs(4) Then Bloque(2) = True
    If rs(6) Then Bloque(3) = True
    If rs(8) Then Bloque(4) = True
    If rs(9) Then Bloque(5) = True
    If rs(10) Then Bloque(6) = True
    If rs(12) Then Bloque(7) = True
    If rs(7) Then Bloque(8) = True

    If rs(3) Then TBloque(1) = rs(18)
    If rs(4) Then TBloque(2) = rs(19)
    If rs(6) Then TBloque(3) = rs(21)
    If rs(8) Then TBloque(4) = rs(23)
    If rs(9) Then TBloque(5) = rs(24)
    If rs(10) Then TBloque(6) = rs(25)
    If rs(12) Then TBloque(7) = rs(27)
    If rs(7) Then TBloque(8) = rs(22)
End If
Call OFV.cerrarconsulta(rs)


strsql = "Select * From Objperfildoc Where codobjeto =" & id

Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
If Not rs.EOF Then

    For bi = 0 To 8
        Bloque(bi) = False
    Next
    For bi = 0 To 8
        TBloque(bi) = " "
    Next
    If rs(3) Then Bloque(1) = True
    If rs(4) Then Bloque(2) = True
    If rs(6) Then Bloque(3) = True
    If rs(8) Then Bloque(4) = True
    If rs(9) Then Bloque(5) = True
    If rs(10) Then Bloque(6) = True
    If rs(12) Then Bloque(7) = True
    If rs(7) Then Bloque(8) = True

    If rs(3) Then TBloque(1) = rs(18)
    If rs(4) Then TBloque(2) = rs(19)
    If rs(6) Then TBloque(3) = rs(21)
    If rs(8) Then TBloque(4) = rs(23)
    If rs(9) Then TBloque(5) = rs(24)
    If rs(10) Then TBloque(6) = rs(25)
    If rs(12) Then TBloque(7) = rs(27)
    If rs(7) Then TBloque(8) = rs(22)

End If

Call OFV.cerrarconsulta(rs)

  '  Bloque(1) = 0
  '  Bloque(2) = 0
  '  Bloque(3) = 1
  '  Bloque(4) = 0
  '  Bloque(5) = 0
  '  Bloque(6) = 0
  '  Bloque(7) = 0
  '  Bloque(8) = 0




sBloque1 = ""
sHTML = ""

   ' Caratula

sBloque1 = sBloque1 & sHead1 & "<BR>"

Call OFV.cerrarconn(cn)
Set cn = OFV.conectar(OFV.strconn0)

strsql = "Select * from usuarios Where ID = '" & sAutor & "' "
Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
If Not rs.EOF Then
    sAutorDes = rs(0) & " - " & rs(1)
End If
Call OFV.cerrarconsulta(rs)
Call OFV.cerrarconn(cn)

sHead1 = OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")

Celda = OFV.GenCelda("", "iett", "", "10%", "", "", "", "Autor: ", "si")
Celda = Celda & OFV.GenCelda("", "ietx", "", "90%", "", "", "", sAutorDes, "si")
sHead1 = sHead1 & OFV.GenRow("", "", "", "", "", "", "", Celda, "si") & OFV.GenfTabla()

sBloque1 = sBloque1 & sHead1 & "<BR>"

If Bloque(2) Then
    Set cn = OFV.conectar(OFV.strconn2)
    strsql = "Select descri from csxobjsector Where codobjeto = " & id & " order by descri "
    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
    If Not rs.EOF Then
        sHead1 = OFV.GenTabla("", "aaa border style='border-style=solid;border-width:1;border-color:#cccccc;'", "", "600", "", "0", "0", "0", "0", "", "", "", "no")
        Celda = OFV.GenCelda("", "istt style='border-color:#cccccc;'", "center", "", "", "", "", TBloque(2), "si")
        sHead1 = sHead1 & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
        Do Until rs.EOF
            Celda = OFV.GenCelda("", "istx style='border-color:#cccccc;'", "", "", "", "", "", rs(0), "si")
            sHead1 = sHead1 & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
            rs.movenext
        Loop
        sHead1 = sHead1 & OFV.GenfTabla()
        sBloque1 = sBloque1 & sHead1 & "<BR>"
    End If
    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)
End If


'Dim sAutId
'Dim sAutDes
'Dim sAutFec
'Dim sAutConf
Dim Sdescripcion, STitulox, Smuestra
sHead1 = ""
xind = 0
Set cn = OFV.conectar(OFV.strconn2)
strsql = "Select * from objwfldet Where codobjeto = " & id & " order by orden, usr "
Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
If Not rs.EOF Then
    sHead1 = OFV.GenTabla("", "aaa border style='border-style=solid;border-width:1;border-color:#cccccc;'", "", "600", "", "0", "0", "0", "0", "", "", "", "no")
    Do Until rs.EOF
        ReDim Preserve sAutId(xind)
        ReDim Preserve sAutDes(xind)
        ReDim Preserve sAutFec(xind)
        ReDim Preserve sAutConf(xind)

        sAutId(xind) = UCase(rs(4))
        sAutDes(xind) = "Desconocido"
        sAutFec(xind) = rs(9)
        If Len(rs(9)) <= 5 Then sAutFec(xind) = "&nbsp;"
        If UCase(rs(8)) = "SC" Then
            sAutConf(xind) = "Sin Confirmar"
            sAutFec(xind) = "&nbsp;"
        ElseIf UCase(rs(8)) = "SI" Then
            sAutConf(xind) = "Confirmado"
        Else
            sAutConf(xind) = "Rechazado"
        End If

        rs.movenext
        xind = xind + 1
    Loop
End If
Call OFV.cerrarconsulta(rs)
Call OFV.cerrarconn(cn)

If sHead1 <> "" and  session("GENAUTOR")  = "S" Then

    Set cn = OFV.conectar(OFV.strconn0)

    Celda = OFV.GenCelda("", "istt colspan=4 style='border-color:#cccccc;'", "center", "", "", "", "", "Lista de Autorizantes", "si")
    sHead1 = sHead1 & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    Celda = OFV.GenCelda("", "istt style='border-color:#cccccc;'", "", "50%", "", "", "", "Autorizante", "si")
    Celda = Celda & OFV.GenCelda("", "istt style='border-color:#cccccc;'", "", "10%", "", "", "", "Fecha", "si")
    Celda = Celda & OFV.GenCelda("", "istt style='border-color:#cccccc;'", "", "10%", "", "", "", "Estado", "si")
    Celda = Celda & OFV.GenCelda("", "istt style='border-color:#cccccc;'", "", "30%", "", "", "", "Firma", "si")
    sHead1 = sHead1 & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

    For xind = 0 To UBound(sAutId)
        strsql = "Select * from usuarios Where ID = '" & sAutId(xind) & "' "
        Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
        If Not rs.EOF Then
            sAutDes(xind) = rs(1)
        End If
        Call OFV.cerrarconsulta(rs)
        Celda = OFV.GenCelda("", "istx style='border-color:#cccccc;'", "", "50%", "30", "", "", sAutDes(xind), "si")
        Celda = Celda & OFV.GenCelda("", "istx style='border-color:#cccccc;'", "", "10%", "30", "", "", sAutFec(xind), "si")
        Celda = Celda & OFV.GenCelda("", "istx style='border-color:#cccccc;'", "", "10%", "30", "", "", sAutConf(xind), "si")
        Celda = Celda & OFV.GenCelda("", "istx style='border-color:#cccccc;'", "", "30%", "30", "", "", "&nbsp;", "si")
        sHead1 = sHead1 & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
    Next
    Call OFV.cerrarconn(cn)

    sHead1 = sHead1 & OFV.GenfTabla()
    sBloque1 = sBloque1 & sHead1 & "<BR>"
End If

sBloque1 = sBloque1 & "<div>&nbsp;</div>"

Call OFV.cerrarconn(cn)
Set cn = OFV.conectar(OFV.strconn2)


  ''''''''''''''''''''''''''''''''
   'cuerpo de descripcion

If Bloque(3) Or Bloque(7) Then

  ' Dim TAnex
  ' Dim TAnexi
  ' Dim TAnexd
  ' Dim TAnexc

   ReDim Preserve TAnex(0)
   ReDim Preserve TAnexi(0)
   ReDim Preserve TAnexd(0)
   ReDim Preserve TAnexc(0)

   oc = 0

  ''14320    strsql = "SELECT DISTINCT COMPLEMENTOS.codcomp, COMPLEMENTOS.codiden, "
  ''14330    strsql = strsql & "COMPLEMENTOS.descri FROM COMPLEMENTOS, "
  ''14340    strsql = strsql & " OBJACCION where COMPLEMENTOS.codcomp = OBJACCION.codcomp "
  ''14350    strsql = strsql & " and OBJACCION.codobjeto =" & id
  ''14360    strsql = strsql & " ORDER BY COMPLEMENTOS.codcomp "


   strsql = ""
   strsql = strsql & " SELECT DISTINCT COMPLEMENTOS.codcomp, COMPLEMENTOS.codiden, "
   strsql = strsql & " COMPLEMENTOS.descri FROM COMPLEMENTOS, OBJACCION "
   strsql = strsql & " where "
   strsql = strsql & " (COMPLEMENTOS.codcomp = OBJACCION.codcomp  and OBJACCION.codobjeto = " & id & " ) "
   strsql = strsql & " UNION "
   strsql = strsql & " ( "
   strsql = strsql & " SELECT DISTINCT COMPLEMENTOS.codcomp, COMPLEMENTOS.codiden, "
   strsql = strsql & " COMPLEMENTOS.descri FROM COMPLEMENTOS, OBJESTDET "
   strsql = strsql & " where "
   strsql = strsql & " (OBJESTDET.dettipo = 'J' AND COMPLEMENTOS.codcomp = OBJESTDET.detimagen "
   strsql = strsql & " and OBJESTDET.codobjeto = " & id & ")) "
   strsql = strsql & " ORDER BY COMPLEMENTOS.codcomp "



   Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
   Do Until rs.EOF
       ocd = oc + 1
       ReDim Preserve TAnex(oc)
       ReDim Preserve TAnexi(oc)
       ReDim Preserve TAnexd(oc)
       ReDim Preserve TAnexc(oc)
       TAnex(oc) = rs(0)
       TAnexi(oc) = "Anexo&nbsp;" & ocd
       TAnexd(oc) = rs(1) & "&nbsp;-&nbsp;" & rs(2)
       TAnexc(oc) = "N"
       oc = oc + 1
       rs.movenext
   Loop
   Call OFV.cerrarconsulta(rs)

   z = 0
   'Dim TAnexa
   'Dim Tanexe
   'Dim TAnexf
   For i = 0 To UBound(TAnex)
       If TAnex(i) <> "" Then
           strsql = "Select * From Compimag Where codcomp = " & TAnex(i) & " Order by secuencia"
           Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
           If Not rs.EOF Then
               TAnexc(i) = "s"
               Do Until rs.EOF
                   ReDim Preserve TAnexa(z)
                   ReDim Preserve Tanexe(z)
                   ReDim Preserve TAnexf(z)
                   TAnexa(z) = rs(3)
                   Tanexe(z) = TAnex(i)
                   TAnexf(z) = "&nbsp;Cuerpo: " & rs(5) & "&nbsp;Descripcion: " & rs(4)
                   z = z + 1
                   rs.movenext
               Loop
               Call OFV.cerrarconsulta(rs)
           Else
               Call OFV.cerrarconsulta(rs)

               strsql = "Select Compimag.*, complementos.codiden, complementos.descri "
               strsql = strsql & " From Compcompuesto, Compimag, complementos where "
               strsql = strsql & "Compcompuesto.compasoc = Compimag.codcomp and "
               strsql = strsql & " Compimag.codcomp=complementos.codcomp "
               strsql = strsql & " and Compcompuesto.compmaster = " & TAnex(i)
               Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
               If Not rs.EOF Then
                   TAnexc(i) = "s"
                   Do Until rs.EOF
                       ReDim Preserve TAnexa(z)
                       ReDim Preserve Tanexe(z)
                       ReDim Preserve TAnexf(z)
                       TAnexa(z) = rs(3)
                       Tanexe(z) = TAnex(i)
                       TAnexf(z) = "&nbsp;" & rs(6) & "&nbsp;" & rs(7) & "&nbsp;Cuerpo: " & rs(5) & "&nbsp;Descripcion: " & rs(4)
                       z = z + 1
                       rs.movenext
                   Loop
               End If
               Call OFV.cerrarconsulta(rs)
           End If
       End If
   Next

  End If
  ''''''''''''''''''''''''''''''''






   'cuerpo de descripcion

If Bloque(8) Then
    strsql = "Select muestra,titulo,descri From Objdscabr Where codobjeto = " & id
    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
    If Not rs.EOF Then
        Sdescripcion = "" & rs("descri")' & chr(13)
        STitulox = "" & rs("titulo")
        Smuestra = rs("muestra")
    Else
        Sdescripcion = "&nbsp;"
        STitulox = "&nbsp;"
        Smuestra = "NO"
    End If
    If Smuestra = "SI" Then
        Celda = OFV.GenCelda("", "ibtt", "", "", "", "", "", Replace(STitulox, Chr(32), "&nbsp;"), "si")
        sh = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
        Sdescripcion = Sdescripcion & Chr(13)
        Sdescripcion = Replace(Sdescripcion, Chr(13), "&nbsp;<br>")
        StrArray = Split(Sdescripcion, "<br>")

        For IArray = 0 To UBound(StrArray) - 1
            Celda = OFV.GenCelda("", "ictx", "", "", "", "", "", Replace(StrArray(IArray), "<br>", ""), "si")
            sh = sh & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
        Next
        sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", sh, "si")
    End If
    Call OFV.cerrarconsulta(rs)
    sBloque1 = sBloque1 & sHTML & "<BR>"
    sHTML = ""

End If
   'Cuerpo de Texto
If Bloque(1) Then
 Dim Strsqlz 
 Dim TObjetivoz 
 Dim sObjetivoz 
 Dim sPoint 
 Dim stexto1 
 Dim Celdat 
 Dim filat 
 Dim pmk 
 Dim sHTMLZ 
 Dim h1 

 Dim XNVER 
 Dim XNSEC 
 Dim sfidk1 
 Strsqlz = "Select * From OBJNWEDT Where codobjeto=0" & id & " Order By consecuen"
 
 Set rsz = OFV.crearconsultaEx(Strsqlz, cn, 1, "")
     
     
 XNVER = "N"
 XNSEC = 0
   

 Do Until rsz.EOF
   
    TObjetivoz = rsz(3)
   
    If IsNull(TObjetivo) Then TObjetivo = " "
   
    TObjetivoz = "<BIG>" & Replace(TObjetivoz, " ", "&nbsp;") & "</BIG>"
   

   
    sObjetivo = ""
    
   
    sObjetivoz = getReadfile(rsz(5) & ".CSE", "")
    If sObjetivoz <> "" Then
         If InStr(sObjetivoz, "<P>") = 0 Then
            sObjetivoz = "<P>" & sObjetivoz & "</P>"
         End If
             
             
             
        Celda = OFV.GenCelda("", "istt ", "", "", "", "", "", TObjetivoz , "si")
       sHTMLZ = sHTMLZ & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
        'Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
        'sHTMLZ = sHTMLZ & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
        Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
        sHTMLZ = sHTMLZ & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
        sObjetivoz = Replace(sObjetivoz, "<P>", "<TR><TD class=uti  width=100% > " )
        sObjetivoz = Replace(sObjetivoz, "</P>", "</TD></TR>")
       'Celda = OFV.GenCelda("", "uti", "", "100%", "", "", "", Replace(sObjetivoz, "&nbsp;", " ", 1, -1, 1), "si")
       sHTMLZ = sHTMLZ & sObjetivoz 'sHTMLZ & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
   
        sHTML = sHTML & OFV.GenTabla("", "", "", "600 ", "", "0", "0", "0", "0", "", "", sHTMLZ, "si")
 End If
   
   
       sHTMLZ = ""
   rsz.movenext
 Loop
   
'6510  sHTMLZ = sHTMLZ & ""
'6520   sHTML = sHTML & OFV.GenfTabla()
   
   
   
    Call OFV.cerrarconsulta(rsz)


   sBloque1 = sBloque1 & sHTML & "<BR>"
   sHTML = ""
   End If

   ' Cuerpo de Contenido
   If Bloque(7) Then
   Fila = ""
   strsql = "Select * From Objestr Where codobjeto = " & id & " order by consecuen"
   Set rsx = OFV.crearconsultaEx(strsql, cn, 1, "")
   Do Until rsx.EOF
       ReDim pdef(13, 18)

       strsql = "Select * From SysEstprty"
       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
       t = 0
       Do Until rs.EOF
           For ti = 1 To 18
               If IsNull(rs(ti)) Then
                   sti = ""
               Else
                   sti = rs(ti)
               End If

               pdef(t, ti - 1) = sti
           Next
           rs.movenext
           t = t + 1
       Loop
       Call OFV.cerrarconsulta(rs)

       EID = rsx(0)
       bsecuen = rsx(4)
       bDescri = rsx(3)

       strsql = "Select * From ObjEstPrty where detcon = " & EID
       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
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
               If IsNull(rs(ti)) Then
                   sti = ""
               Else
                   sti = rs(ti)
               End If

               pdef(t, ti - 5) = sti
           Next
           rs.movenext
       Loop
       Call OFV.cerrarconsulta(rs)

       PInden = pdef(0, 7)
       If PInden <> "NO" Then PInden = "SI"

       Pnumer = pdef(0, 12)
       Pnumcls = pdef(0, 14)
       ReDim pnumtip(13)
       For i = 0 To 13
           pnumtip(i) = pdef(i, 13)
           If pnumtip(i) = "" Then pnumtip(i) = "1"
       Next

           'Estructura
       Celda = OFV.GenCelda("", "ibtt", "", "60%", "", "", "", bDescri, "si")
       tabla1 = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
       tabla1 = OFV.GenTabla("", "", "", "600", "", "0", "", "0", "0", "", "", tabla1, "si")
       Fila = Fila & tabla1 & OFV.GenTabla("CB", "aaa  ", "", "600", "", "0", "", "0", "0", "", "", "", "no")

       Celda = ""
       For i = 1 To 10
           Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
       Next
       Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

       strsql = "SELECT * from Objestdet where detcon = " & EID & " and dettipo = 'A' "
       strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
       strsql = strsql & " order by detsecuen"
       'response.write strsql
       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

       sant = "X"
       bseci = 0
       bsecix = 0
       ReDim bsec(8)
       ReDim Bnum(8)
       ReDim Bnumg(8)
       For i = 0 To 7
           bsec(i) = 0
           Bnum(i) = 0
           Bnumg(i) = 0
       Next

       Do While Not rs.EOF
           Do While Not rs.EOF
               sclave = rs(0)
               stip = rs(5)
               smodi = rs(29)
               stipmodi = rs(30)
               sForm = "edestrdet" & sclave

               sImage = "&nbsp;"
               If bVersion = "SI" And smodi Then
                   If stipmodi = "M" Then
                       sImage = OFV.GenImage("IC", "", "", "", "Imagenes/Botmodx.jpg", "0", "")
                   ElseIf stipmodi = "A" Then
                       sImage = OFV.GenImage("IC", "", "", "", "Imagenes/Botagrx.jpg", "0", "")
                   End If
               End If
               Celda = OFV.GenCelda("", "ictx", "", "20", "", "", "", sImage, "si")

                   'Secuencia
               Select Case stip
                       Case "A"
                       t = 7
                       bseci = 0
                       bsecix = 0
                       bsec(1) = rs(4)
                       Bnum(0) = Bnum(0) + 1
                       Bnumg(0) = Bnumg(0) + 1
                       For i = bseci + 1 To 8
                           Bnum(i) = 0
                           Bnumg(i) = 0
                       Next
                   Case "B"
                       t = 8
                       bseci = 1
                       bsecix = 1
                       bsec(2) = rs(4)
                       Bnum(1) = Bnum(1) + 1
                       Bnumg(1) = Bnumg(1) + 1
                       For i = bseci + 1 To 8
                           Bnumg(i) = 0
                           Bnum(i) = 0
                       Next
                   Case "C"
                       t = 9
                       bseci = 2
                       bsecix = 2
                       bsec(3) = rs(4)
                       Bnumg(2) = Bnumg(2) + 1
                       Bnum(2) = Bnum(2) + 1
                       For i = bseci + 1 To 8
                           Bnumg(i) = 0
                           Bnum(i) = 0
                       Next
                   Case "D"
                       t = 10
                       bseci = 3
                       bsecix = 3
                       bsec(4) = rs(4)
                       Bnumg(3) = Bnumg(3) + 1
                       Bnum(3) = Bnum(3) + 1
                       For i = bseci + 1 To 8
                           Bnumg(i) = 0
                           Bnum(i) = 0
                       Next
                   Case "E"
                       t = 11
                       bseci = 4
                       bsecix = 4
                       bsec(5) = rs(4)
                       Bnum(4) = Bnum(4) + 1
                       Bnumg(4) = Bnumg(4) + 1
                       For i = bseci + 1 To 8
                           Bnumg(i) = 0
                           Bnum(i) = 0
                       Next
                   Case "F"
                       t = 12
                       bseci = 5
                       bsecix = 5
                       bsec(6) = rs(4)
                       Bnum(5) = Bnum(5) + 1
                       Bnumg(5) = Bnumg(5) + 1
                       For i = bseci + 1 To 8
                           Bnumg(i) = 0
                           Bnum(i) = 0
                       Next
                   Case "G"
                       t = 13
                       Select Case sant
                               Case "A"
                               ta = 7
                               bseci = 0
                           Case "B"
                               ta = 8
                               bseci = 1
                           Case "C"
                               ta = 9
                               bseci = 2
                           Case "D"
                               ta = 10
                               bseci = 3
                           Case "E"
                               ta = 11
                               bseci = 4
                           Case "F"
                               ta = 12
                               bseci = 5
                       End Select
                       bsecix = bseci + 1
                       bsec(7) = rs(4)
                       Bnum(6) = Bnum(6) + 1
                       Bnumg(6) = Bnumg(6) + 1
                   Case "H"
                       Select Case sant
                               Case "A"
                               bseci = 0
                           Case "B"
                               bseci = 1
                           Case "C"
                               bseci = 2
                           Case "D"
                               bseci = 3
                           Case "E"
                               bseci = 4
                           Case "F"
                               bseci = 5
                       End Select
                       bsecix = bseci + 1
                       bsec(7) = rs(4)
                       Bnum(7) = Bnum(7) + 1
                       Bnumg(7) = Bnumg(7) + 1
                   Case "I"
                       Select Case sant
                               Case "A"
                               bseci = 0
                           Case "B"
                               bseci = 1
                           Case "C"
                               bseci = 2
                           Case "D"
                               bseci = 3
                           Case "E"
                               bseci = 4
                           Case "F"
                               bseci = 5
                       End Select
                       bsecix = bseci + 1
                       bsec(7) = rs(4)
                       Bnum(8) = Bnum(8) + 1
                       Bnumg(8) = Bnumg(8) + 1
                   Case "J"
                       Select Case sant
                               Case "A"
                               bseci = 0
                           Case "B"
                               bseci = 1
                           Case "C"
                               bseci = 2
                           Case "D"
                               bseci = 3
                           Case "E"
                               bseci = 4
                           Case "F"
                               bseci = 5
                       End Select
                       bsecix = bseci + 1
                       bsec(7) = rs(4)
                       Bnum(8) = Bnum(8) + 1
                       Bnumg(8) = Bnumg(8) + 1

               End Select

               If stip <> "H" And stip <> "I" And stip <> "J" Then
                   If rs(31) = "NO" Then
                       For ti = 11 To 28
                           If IsNull(rs(ti)) Then
                               sti = ""
                           Else
                               sti = rs(ti)
                           End If

                           pdef(t, ti - 11) = sti

                       Next
                       pdef(t, 18) = "NO"
                   Else
                       For ti = 0 To 17
                           pdef(t, ti) = pdef(t - 7, ti)
                       Next
                       pdef(t, 18) = "NO"
                   End If
               End If

               BsecD = ""
               For i = 1 To bsecix
                   BsecD = BsecD & bsec(i) & "."
               Next

               Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")

               If PInden = "SI" Then
                   If bsecix = 7 Then
                       bspan = ""
                   Else
                       bspan = "aaa colspan=" & 7 - bsecix
                   End If
                   For i = 1 To bsecix
                       Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                   Next
               Else
                   If stip <> "G" And stip <> "H" And stip <> "I" And stip <> "J" Then
                       bspan = "aaa colspan=7"
                   Else
                       bspan = "aaa colspan=6"
                       Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                   End If
               End If

               strsql = "SELECT * from Objestdet where detcon = " & EID & " and detdepen = " & sclave
               strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
               strsql = strsql & " order by detsecuen"

               If pdef(t, 18) = "NO" Then
                   tx = t
               Else
                   tx = t - 7
               End If


               Pnumcls = pdef(tx, 14)

               If Pnumcls = "" Then Pnumcls = "Simple"



               If stip <> "H" And stip <> "I" And stip <> "J" Then
                   pnumtip(tx) = pdef(tx, 13)
                   If pnumtip(tx) = "" Then pnumtip(tx) = "1"
               End If

               If stip = "G" Then
                   If pdef(tx, 9) = "NO" Then
                       mpvin = "NO"
                   Else
                       mpvin = "SI"
                   End If
               End If

               If pnumtip(tx) = "NO" Then Pnumcls = pnumtip(tx)

               pfont = pdef(tx, 0)
               pfsize = pdef(tx, 1)
               pfcol = pdef(tx, 2)
               pbgcol = pdef(tx, 3)
               PBorde = pdef(tx, 4)
               PFtipo = pdef(tx, 5)

               PSang = pdef(tx, 6)
               txi = tx
               If tx > 6 Then txi = tx - 7
               PEspa = pdef(txi, 8)

               Palign = pdef(tx, 16)
               If Palign = "" Then Palign = "left"
               Ptalig = pdef(tx, 17)

               pborst = ""
               If PBorde <> "" And PBorde <> "NO" Then
                   PBCol = ""
                   If UCase(PBorde) = "SOLID" Then PBCol = ";border-color:#000000"
                   pborst = "border-width:thin;border-style:" & PBorde & PBCol & ";"
               End If

               PFtipoS = ""
               PFtipoF = ""
               If PFtipo <> "" Then
                   For vt = 1 To Len(PFtipo)
                       Select Case Mid(PFtipo, vt, 1)
                                 Case "1"
                               PFtipoS = PFtipoS & "<B>"
                               PFtipoF = PFtipoF & "</B>"
                           Case "2"
                               PFtipoS = PFtipoS & "<I>"
                               PFtipoF = PFtipoF & "</I>"
                           Case "4"
                               PFtipoS = PFtipoS & "<U>"
                               PFtipoF = PFtipoF & "</U>"
                           Case "8"
                               PFtipoS = PFtipoS & "<S>"
                               PFtipoF = PFtipoF & "</S>"
                       End Select
                   Next
               End If
               pStyle = PFtipoS
               pStyleF = PFtipoF
               nfonta = "<font face=" & Chr(34) & pfont & Chr(34) & " >"
               nfontx = "</font>"

               pfonst = "font-family:'" & Replace(pfont, " ", "-") & "';color:" & pfcol & ";font-size:"
               pfonst = pfonst & pfsize & "pt;"

               bspan = bspan & " STYLE={" & pborst & pfonst & "} "

                     ' TITULOS Y SUBTITULOS --------------------------------------------------
               If stip <> "G" And stip <> "H" And stip <> "I" And stip <> "J" Then
                         'Descripcion
                   PnumS = ""
                   If Pnumer = "SI" Then
                       Select Case UCase(Pnumcls)
                                 Case "SIMPLE"
                               Select Case pnumtip(tx)
                                         Case "1"
                                       BnumS = Bnumg(bseci)
                                   Case Else
                                       BnumS = ConvNumer(Bnumg(bseci), pnumtip(tx))
                               End Select

                               PnumS = BnumS & ". "
                           Case "COMPUESTA"
                               For i = 0 To bseci
                                   Select Case pnumtip(i + 7)
                                             Case "NO"
                                           BnumS = ""
                                       Case "1"
                                           BnumS = Bnumg(i) & "."
                                       Case Else
                                           BnumS = ConvNumer(Bnumg(i), pnumtip(i + 7)) & "."
                                   End Select

                                   PnumS = PnumS & BnumS
                               Next
                               PnumS = PnumS & " "
                       End Select
                   End If
                   Celda = Celda & OFV.GenCelda("", bspan, Palign, "", "", pbgcol, "", pStyle & nfonta & PnumS & rs(7) & nfontx & pStyleF, "si")
                   Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                   Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

               End If

                      ' PARRAFOS DE TEXTO    --------------------------------------------------
               If stip = "G" Then
                   PnumS = ""
                   If Pnumer = "SI" Then
                       Select Case UCase(Pnumcls)
                                  Case "NO"
                               Bnumg(6) = Bnumg(6) - 1
                           Case "SIMPLE"
                               Select Case pnumtip(tx)
                                          Case "1"
                                       BnumS = Bnumg(6)
                                   Case Else
                                       BnumS = ConvNumer(Bnumg(6), pnumtip(tx))
                               End Select

                               PnumS = BnumS & ". "
                           Case "COMPUESTA"
                               For i = 0 To 6
                                   If Bnumg(i) > 0 Then
                                       Select Case pnumtip(i + 7)
                                                  Case "NO"
                                               BnumS = ""
                                           Case "1"
                                               BnumS = Bnumg(i) & "."
                                           Case Else
                                               BnumS = ConvNumer(Bnumg(i), pnumtip(i + 7)) & "."
                                       End Select

                                       PnumS = PnumS & BnumS
                                   End If
                               Next
                               PnumS = PnumS & " "
                       End Select
                   End If

                   tx = ta
                   If Ptalig = "" Then
                       Ptalig = pdef(tx, 17)
                       If Ptalig = "" Then Ptalig = "left"
                   End If

                   If PSang = "" Then PSang = 0
                   PsangS = ""
                   For Vs = 1 To PSang
                       PsangS = PsangS & "&nbsp;"
                   Next
                   If pdef(tx, 9) <> "SI" Then
                       PVin = ""
                   Else
                       PVFont = pdef(tx, 15)
                       PVtipo = pdef(tx, 10)
                       pVind = pdef(tx, 11)
                       If pVind = "" Then pVind = 0
                       PVindS = ""
                       For Vi = 1 To pVind
                           PVindS = PVindS & "&nbsp;"
                       Next
                       PVindS = "<font face=tahoma size=1>" & PVindS & "</font>"
                       PVin = PFtipoS
                       PVin = PVin & PVtipo & PFtipoS & PVindS

                       pfonst = "aaa valign=TOP Style={"
                       pfonst = pfonst & "font-family:" & PVFont & ";color:" & pfcol & ";font-size:"
                       pfonst = pfonst & pfsize & "pt;} "
                   End If

                   If mpvin <> "SI" Then
                       PVin = ""
                   End If

                          'Texto

                   wrtexto = rs(8) & ""

                   XDescri = ""

                   If Len(rs("rellink") & "") > 0 And rs("rellink") <> "0" And rs("rellink") <> " " Then
                       Wlinkd = CLng(rs("rellink"))
                       strsql = "SELECT *  From objlnk where objid = " & Wlinkd
                              '                 wrtexto = wrtexto & "<br>1<br>" & strsql & "<br>"
                       Set rst = OFV.crearconsultaEx(strsql, cn, 1, "")
                       If Not rst.EOF Then
                           codobjetoL = CLng(rst("codobjetoL"))
                                  ' wtipoLc = rst(7)
                                  '                        wrtexto = wrtexto & "<br>2<br>" & codobjetoL & "<br>Ltipo=" & rst("Ltipo") & "<br>"
                           If rst("Ltipo") = "L" Then
                               strsql = "Select * From csGenObjetos where codobjeto=" & codobjetoL
                                      '                                   wrtexto = wrtexto & "<br>3x<br>" & strsql & "<br>"
                               Set rst2 = OFV.crearconsultaEx(strsql, cn, 1, "")
                               XDescri = " <u><b><i>" & rst2("xdescri") & "</i></b></u> "
                               Call OFV.cerrarconsulta(rst2)
                           Else
                               If Not rst.EOF Then
                                   XDescri = " <u><b><i>" & rst("LurlDesc") & "</i></b></u> "
                               End If
                           End If


                       End If
                       Call OFV.cerrarconsulta(rst)
                   End If



                   wrtexto = wrtexto & Chr(13) & XDescri & Chr(13)


                   If Len(wrtexto) > 0 Then


                       wrtexto = Replace(wrtexto, Chr(13), "<br>")
                       irtextoO = CountOccurrences(wrtexto, "<br>", vbTextCompare)


                              '            Descri = ""
                              '            wrtexto = wrtexto & "<br>11<br>" & rs("detimagen") & " wwww<br>" & rs(0) & "<br>---"
                              '           wrtexto = wrtexto & " <br><i>" & Descri & "</i>"

                              ' Celda = Celda & OFV.GenCelda("", bspan, Ptalig & " valign=TOP", "", "", pbgcol, "", pStyle & PsangS & sTextoV & pStyleF, "si")
                       If irtextoO = 0 Then

                           If PnumS <> "" Then
  '12710                                Celda = Celda & OFV.GenCelda("", "aaa valign=top STYLE={" & pfonst & "} ", "", "1%", "", "", "", pStyle & PnumS & wrtexto & pStyleF, "si")
                               Celda = Celda & OFV.GenCelda("", "aaa valign=top STYLE={" & pfonst & "} ", "", "1%", "", "", "", pStyle & PnumS & pStyleF, "si")
                           End If

                           If PVin <> "" Then
                               Celda = Celda & OFV.GenCelda("", pfonst, "", "1%", "", "", "", PVin, "si")
                           End If

                           Celda = Celda & OFV.GenCelda("", bspan, Ptalig & " valign=top", "", "", pbgcol, "", pStyle & nfonta & PsangS & wrtexto & nfontx & pStyleF, "si")



                           If PnumS <> "" Then
                               Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                           End If

                           If PVin = "" Then
                               Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                           End If
                           Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


                       Else

                           For irtextoI = 1 To irtextoO

                               irtextoc = InStr(1, wrtexto, "<br>")

                               wrtextosplit = Mid(wrtexto, 1, irtextoc + 3)
                               wrtextosplit = Replace(wrtextosplit, "<br>", "")
                               If Len(wrtextosplit) = 0 Then wrtextosplit = " "
                               If Len(wrtexto) > (irtextoc + 4) Then
                                   wrtexto = Mid(wrtexto, irtextoc + 4, Len(wrtexto))
                               Else
                                   wrtexto = " "
                               End If

                               If irtextoI = 1 Then

                                   If PnumS <> "" Then
                                       Celda = Celda & OFV.GenCelda("", "aaa valign=top STYLE={" & pfonst & "} ", "", "1%", "", "", "", pStyle & nfonta & PnumS & nfontx & pStyleF, "si")
                                   End If

                                   If PVin <> "" Then
                                       Celda = Celda & OFV.GenCelda("", pfonst, "", "1%", "", "", "", PVin, "si")
                                   End If

                                   Celda = Celda & OFV.GenCelda("", bspan, Ptalig & " valign=top", "", "", pbgcol, "", pStyle & nfonta & PsangS & wrtextosplit & nfontx & pStyleF, "si")

                                   If PnumS <> "" Then
                                       Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                                   End If

                                   If PVin = "" Then
                                       Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                                   End If
                                   Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


                               Else

                                          ' 1ra celda
                                   sImage = "&nbsp;"
                                   Celda = OFV.GenCelda("", "ictx", "", "20", "", "", "", sImage, "si")

                                          ' 2nda celda

                                   Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")

                                   If PInden = "SI" Then
                                       For i = 1 To bsecix
                                           Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                                       Next
                                   Else
                                       Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                                   End If

                                          ' 3ra linea
                                          'Texto

                                   If PnumS <> "" Then
                                       Celda = Celda & OFV.GenCelda("", "aaa STYLE={" & pfonst & "} ", "", "1%", "", "", "", "&nbsp;", "si")
                                   End If

                                   If PVin <> "" Then
                                       Celda = Celda & OFV.GenCelda("", pfonst, "", "1%", "", "", "", "&nbsp;", "si")
                                   End If

                                   Celda = Celda & OFV.GenCelda("", bspan, Ptalig & " valign=top", "", "", pbgcol, "", pStyle & nfonta & PsangS & wrtextosplit & nfontx & pStyleF, "si")

                                   If PnumS <> "" Then
                                       Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                                   End If

                                   If PVin = "" Then
                                       Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                                   End If
                                   Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

                               End If
                           Next
                       End If



                   Else

                       If PnumS <> "" Then
                           Celda = Celda & OFV.GenCelda("", "aaa valign=top STYLE={" & pfonst & "} ", "", "1%", "", "", "", pStyle & PnumS & pStyleF, "si")
                       End If

                       If PVin <> "" Then
                           Celda = Celda & OFV.GenCelda("", pfonst, "", "1%", "", "", "", PVin, "si")
                       End If

                       Celda = Celda & OFV.GenCelda("", bspan, Ptalig & " valign=top", "", "", pbgcol, "", pStyle & PsangS & "&nbsp;" & pStyleF, "si")

                       If PnumS <> "" Then
                           Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                       End If

                       If PVin = "" Then
                           Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                       End If
                       Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


                   End If



               End If

                      ' IMAGENES            --------------------------------------------------
               If stip = "H" Then

                          'Descripcion
                   sreduc = ""
                   sbord = ""
                   salng = ""

                   If rs(12) <> "" Then sreduc = " width=" & rs(12) & "% "
                   If rs(15) <> "" Then
                       If rs(15) = "SI" Then sbord = " border=1 "
                   End If
                   If rs(16) <> "" Then
                       If rs(16) = "CENTRADO" Then salng = " center"
                       If rs(16) = "DERECHA" Then salng = " right"
                       If rs(16) = "IZQUIERDA" Then salng = " left"
                   End If

                   If rs(9) <> "" Then
                       sLink = "<img id=IMD src=" & Chr(34) & "estructura/" & rs(9) & Chr(34) & " "
                       sLink = sLink & sreduc & sbord & " >"
                   Else
                       sLink = ""
                   End If
                   Celda = Celda & OFV.GenCelda("", bspan, salng, "", "", "", "", sLink, "si")
                   Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                   Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

               End If

                      ' Tablas            --------------------------------------------------
               If stip = "I" Then

                   sLink = "&nbsp;"

                   Celda = Celda & OFV.GenCelda("", bspan, "", "", "", "", "", sLink, "si")
                   Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                   Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

               End If

                     ' COMPLEMENTOS            --------------------------------------------------

 If stip = "J" Then



  If rs(9) <> "" Then

      sreduc = ""
      sbord = ""
      salng = ""

      strsql = "SELECT   nombre,TIPOCOMP FROM CSACOMPLE where codcomp = 0" & rs(9)
      Set rst = OFV.crearconsultaEx(strsql, cn, 1, "")
      If Not rst.EOF Then
         sreduc = rst(0)
         sbord = rst(1)
      End If
      Call OFV.cerrarconsulta(rst)

      strsql = "SELECT archivo FROM tipocomp where tipo = '" & sbord & "' "
      Set rst = OFV.crearconsultaEx(strsql, cn, 1, "")
      If Not rst.EOF Then
         salng = rst(0)
      End If
      Call OFV.cerrarconsulta(rst)

    sInput = sreduc

    If Len(sInput) > 50 Then
       sInput = Left(sInput, 47) & "..."
    End If

     sCompAn = ""
     For oc = 0 To UBound(TAnex)
         If CLng(TAnex(oc)) = CLng(rs(9)) Then
            sCompAn = "&nbsp;(ver " & TAnexi(oc) & ")&nbsp;"
             ''   TAnexd(oc) = sComplemento
            Exit For
         End If
     Next

     sInput = "&nbsp;&nbsp;<B>" & Replace(sInput, " ", "&nbsp;") & "&nbsp;" & "</b>"
     sInput = sInput & sCompAn


     sLink = "<img id=IC src=" & Chr(34) & "iconos/" & salng & Chr(34) & " "
     sLink = sLink & " >"

     sImage = sLink & sInput


     sLink = "<BR>" & sImage & "<BR>"

 Else

     sLink = "&nbsp;"

 End If


     Celda = Celda & OFV.GenCelda("", bspan, "", "", "", "", "", sLink, "si")
     Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
     Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
 End If

    '------------------------------------------------

               Select Case stip
                          Case "A"
                       sant = "A"
                       rs.movenext
                       Set rsa = rs
                       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
                   Case "B"
                       sant = "B"
                       rs.movenext
                       Set rsb = rs
                       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
                   Case "C"
                       sant = "C"
                       rs.movenext
                       Set rsc = rs
                       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
                   Case "D"
                       sant = "D"
                       rs.movenext
                       Set rsd = rs
                       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
                   Case "E"
                       sant = "E"
                       rs.movenext
                       Set rse = rs
                       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
                   Case "F"
                       sant = "F"
                       rs.movenext
                       Set rsf = rs
                       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
                   Case "G"
                       rs.movenext
                   Case "H"
                       rs.movenext
                   Case "I"
                       rs.movenext
                   Case "J"
                       rs.movenext
               End Select
                      '        Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

               sBloque1 = sBloque1 & Fila
               Fila = ""

               If PEspa = "" Then PEspa = 0
               For y = 1 To PEspa
                   Celda = ""
                   For i = 1 To 10
                       Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
                   Next
                   Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

                   sBloque1 = sBloque1 & Fila
               Next

           Loop

           ssale = "no"
           Do Until ssale = "si"

               If sant = "A" Then
                   Set rs = rsa
                   sant = "X"
               ElseIf sant = "B" Then
                   Set rs = rsb
                   sant = "A"
               ElseIf sant = "C" Then
                   Set rs = rsc
                   sant = "B"
               ElseIf sant = "D" Then
                   Set rs = rsd
                   sant = "C"
               ElseIf sant = "E" Then
                   Set rs = rse
                   sant = "D"
               ElseIf sant = "F" Then
                   Set rs = rsf
                   sant = "E"
               End If

               If (Not rs.EOF) Or sant = "X" Then
                   ssale = "si"
               End If

           Loop

       Loop

       Celda = ""
       For i = 1 To 10
           Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
       Next

       Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

       sBloque1 = sBloque1 & Fila
       Celda = ""
       For i = 1 To 10
           Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
       Next

       Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
       sBloque1 = sBloque1 & Fila

       Call OFV.cerrarconsulta(rs)

       Fila = OFV.GenfTabla()

       rsx.movenext

   Loop

   If Fila <> "" Then sBloque1 = sBloque1 & OFV.GenfTabla()

   Call OFV.cerrarconsulta(rsx)

   End If


      ' Cuerpo de Procedimientos
   If Bloque(3) Then


'''''''''''

'''''''''''''
   sHTML = ""

   sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "-2", "-1", "", "", "", "no")

          'array de acciones del link
   Dim wsmodo()
   Dim wsmodoB()
   Dim wstipo(6)
   Dim ym, tm, ymm
   Dim wtipoL, wUrl, wUrlDesc, wtipoLNK, Sarma
   ReDim wsmodo(0)
   ReDim wsmodoB(0)

   ym = 1

   strsql = "select * from modo "
   strsql = strsql & " Order By tipomodo"
   Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
   If Not rs.EOF Then
       Do Until rs.EOF
           tm = ym
           ReDim Preserve wsmodo(ym)
           ReDim Preserve wsmodoB(ym)
           wsmodo(ym) = rs(0)
           wstipo(0) = rs(2)
           For ymm = 1 To 6
               wstipo(ymm) = rs(ymm + 3)
           Next
           wsmodoB(ym) = wstipo
           ym = ym + 1
           rs.movenext
       Loop
   Else
       tm = ym
       ReDim Preserve wsmodo(ym)
       ReDim Preserve wsmodoB(ym)
       wsmodo(ym) = "NO"
   End If

   Call OFV.cerrarconsulta(rs)
   
   




 
    ssql="select * from CSCEventoSectorPerfil where obcodobjeto=" & id & "  order by secuen, secuencia,codproced "
  '  set rs2w = ofv.crearconsultaEx(ssql,cn,1,parametros)
   ' response.write ssql
    
   Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
   Celda = OFV.GenCelda("", "ibtt", "", "", "", "", "", TBloque(3), "si")
   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si") & OFV.GenfTabla()
   
   
   
   
   sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")
   
  

  set rsdd = ofv.crearconsultaEx(ssql,cn,1,parametros)


''   Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
''   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
''   Celda = OFV.GenCelda("", "ibtt", "", "", "", "", "", "sql: " & ssql & "<br>eof:" & rsdd.eof , "si")
''   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si") & OFV.GenfTabla()

          eveant  = ""

    EVANT = 0
  do while not rsdd.eof
     nIDd   = rsdd(0)
     tit   = rsdd(7)
     frec  = rsdd(8)
     EVSC  = rsdd(2)
     EVPS  = rsdd(6)
     EVDS  = rsdd(4)
     SECD  = rsdd(3)
     PERD  = rsdd(5)
     sEven = ofv.convertircar(sEven," ","_")

    Fila = ""

    sPoint = "EV" & nID



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    EVANT = EVSC
    Celda = ofv.GenCelda(sPoint,"iett","","10%","","","","EVENTO:","si")

    stexto1 = EVSC & "." & EVPS & "&nbsp;-&nbsp;" & EVDS

    Celda = Celda & ofv.GenCelda("","ietx colspan=3 ","","60%","","","",stexto1,"si")

    Fila = FILA & ofv.GenRow("","","","","","","",celda,"si")


    Celda = ofv.GenCelda("","iett","","10%","","","","SECTOR:","si")

    Celda = Celda & ofv.GenCelda("","ietx","","60%","","","",SECD,"si")



    Celda = Celda & ofv.GenCelda("","iett","","10%","","","","FRECUENCIA:","si")

    Strsql = "SELECT descri from frecuencia where codfrec = '" & frec & "' "
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs1.eof then
       wsfrec = rs1(0)
    else
       wsfrec = "SIN&nbsp;DEFINIR"
    end if

    wsfrec = wsfrec

    Celda = Celda & ofv.GenCelda("","ietx","","20%","","","",wsfrec,"si")


    Fila = FILA & ofv.GenRow("","","","","","","",celda,"si")




    Celda = ofv.GenCelda("","iett","","10%","","","","PERFIL&nbsp;:","si")


    Celda = Celda & ofv.GenCelda("","ietx colspan=3 ","","60%","","","",perD,"si")


    Fila = FILA & ofv.GenRow("","","","","","","",celda,"si")



    Strsql = "SELECT referencia from objproced where codproced = '" & rsdd(0) & "' "
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs1.eof then
       wsref = rs1(0) 
       if right(wsref,1)<> chr(13)  then wsref =wsref & chr(13)
    else
       wsref = "NO"
    end if

    if isnull(wsref) then wsref = "NO"


' Referencia del procedimiento

  if wsref <> "NO" then
     sInput = ""
    
     sReferencia = Replace(wsref, Chr(13), "&nbsp;<br>")
     StrArray = Split(sReferencia, "<br>")
     For IArray = 0 To UBound(StrArray) - 1               
      sInput = Replace(StrArray(IArray), "<br>", "")
      Celda = ofv.GenCelda("","ictx style='text-align:justify;' ","","100% colspan=4","","","",sInput,"si")
      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
     next
  end if

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-bottom-style:double;border-bottom-width:1;border-bottom-color:#000088;' ","","600px border=0 ","","0","0","-1","2","","",Fila,"si") & "<BR>"
    sHTML = sHTML & ofv.GenTabla("","aaa ","","600px ","","0","0","-2","-1","","","","no")

  Fila = ""

  bColor = "False"
  if nID <> "no" then

    redim wsmodo(0)
    redim wsmodoB(0)

    ym = 1

    strsql = "select * from modo "
    strsql = strsql & " Order By tipomodo"
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
       do until rs2.eof
          tm = ym
          redim preserve wsmodo(ym)
          redim preserve wsmodoB(ym)
          wsmodo(ym) = rs2(0)
          wstipo(0) = rs2(2)
          for ymm = 1 to 6
              wstipo(ymm) = rs2(ymm + 3)
          next
          wsmodoB(ym) = wstipo
          ym = ym + 1
       rs2.movenext
       loop
    else
          tm = ym
          redim preserve wsmodo(ym)
          redim preserve wsmodoB(ym)
          wsmodo(ym) = "NO"
    end if





    strsql = "select * from objaccion "
    strsql = strsql & "WHERE codproced = '" &  nidd & "'"
    strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
    strsql = strsql & " Order By secuen"
   ' response.write strsql & " 222222 <br>"
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    posicion = 0
    
    do while not rs2.eof

      nAccion   = rs2(0)
      nAccionID = rs2(18)
      xfclas    = "dt"
      sImage    = "&nbsp;"

      if bVersion = "SI" and rs2(14) then
      if rs2(15) = "M" then
         xftitl = "Modificado."
         xfclas = "btbs"
         sglmod = "M" '"<div class='" & xfclas & "' ><font class=btic >4</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      elseif rs2(15) = "A" then
         xftitl = "Agregado."
         xfclas = "btba"
         sglmod = "A" '"<div class='" & xfclas & "' ><font class=btic >6</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      elseif rs2(15) = "E" then
         xftitl = "Eliminado."
         xfclas = "btbn"
         sglmod = "E" '"<div class='" & xfclas & "' ><font class=btic >5</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      end if
      end if



      sPoint = "T" & nAccion


      'Acciones
      if rs2(4) <> "" then
         Strsql = "Select descri,MODO From Tipoaccalt Where tipoaccionalt = " & rs2(4)
         set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if  not rsAcciones.eof then
             sAccionM = rsAcciones(0)
             sModo   = rsAcciones(1)
         end if 
         call ofv.cerrarconsulta(rsAcciones)
      end if


      Celda  = ofv.GenCelda(sPoint,"uts title='" & xftitl & "' ","","1%","","","","<b>" & sImage & "</b>","si")


      tenc = "NO"
      for ym=1 to tm
          if sModo = wsmodo(ym) then
             tenc = "SI"
             for ymm=0 to 6
                 wstipo(ymm) = wsmodoB(ym)(ymm)
             next

             exit for
          elseif sModo < wsmodo(ym) then
             tenc = "NO"
             exit for
          end if
      next

      if ym > tm then  tenc = "NO"

      if   tenc = "NO" then
             for ymm=0 to 6
                 wstipo(ymm) = 0
             next
      end if




   





      sVerbo   = sAccionM
      sAcciont = sAccionM
      i        = instr(sVerbo," ")

      if i > 0 then
        sVerbo   = "<b>" &   left(sAccionM,i-1) & "</b>"
        sAcciont = mid(sAccionM,i+1)
      end if

      sInput = rs2(10) & ".&nbsp;<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & sAcciont
      Celda  = Celda & ofv.GenCelda("","ictt colspan=5  ","","","","","", sInput ,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")





      'Texto de Referencia
      sArma = rs2(11) & ""
      if len(sArma) > 0 then
       if right(sarma,1)<> chr(13)  then sarma =sarma & chr(13)
        sArma = replace(sArma,chr(13),"<br>")
        
        StrArray = Split(sArma, "<br>")
      
        For IArray = 0 To UBound(StrArray) - 1
         Celda =  ofv.GenCelda("","ictx colspan=3","","20","","","","&nbsp;","si")
         Celda = Celda & ofv.GenCelda("","ictx colspan=4 style='text-align:justify;' ","","","","","",Replace(StrArray(IArray), "<br>",""),"si")
         Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
        Next 
      end if

      sArma = rs2(16) & ""
      if len(sArma) > 0 then
       
       if right(sarma,1) <> chr(13)   then sarma =sarma & chr(13)
        sArma = replace(sArma,chr(13),"<br>")
        StrArray = Split(sArma, "<br>")
         For IArray = 0 To UBound(StrArray) - 1 
          Celda = ofv.GenCelda("","ictx colspan=3","","20","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","ictx colspan=4 style='text-align:justify;' ","","","","","",Replace(StrArray(IArray), "<br>",""),"si")
          Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
         Next 
      end if

      sArma = rs2(17) & ""
      if len(sArma) > 0 then
       if right(sarma,1)<> chr(13)  then  sarma =sarma & chr(13)
        sArma = replace(sArma,chr(13),"<br>")
        StrArray = Split(sArma, "<br>")
        For IArray = 0 To UBound(StrArray) - 1  
         Celda = ofv.GenCelda("","ictx colspan=3","","20","","","","&nbsp;","si")
         Celda = Celda & ofv.GenCelda("","ictx colspan=4 style='text-align:justify;' ","","","","","",Replace(StrArray(IArray), "<br>",""),"si")
         Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
        Next 
      end if


''''' impresion de links anulada ''''''''''''''''''''''''
      'link
      if wstipo(3) then
         wlnkptr = "0" & rs2(19)
         if not isnull(wlnkptr) then
          if clng(wlnkptr) > 0 then
           Strsql = "SELECT * From objlnk where objid = " & rs2(19)
           set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsx.eof then
              wtipoL   = rsx(6)
              wtipoLc  = rsx(7) & ""
              wUrl     = rsx(8) & ""
              wUrlDesc = rsx(9) & ""
              wtipoLNK = rsx(10) & ""
              wNV1    = rsX(11) & ""
              wNV2    = rsX(12) & ""
              wNV3    = rsX(13) & ""
              IF  WNV1 = "" THEN WNV1 = "0"
              IF  WNV2 = "" THEN WNV2 = "0"
              IF  WNV3 = "" THEN WNV3 = "0"
           else
              wtipoL   = 0
              wtipoLc  = "NO"
              wUrl     = "NO"
              wUrlDesc = ""
              wtipoLNK = "NO"
              WNV1 = "0"
              WNV2 = "0"
              WNV3 = "0"
           end if
           call ofv.cerrarconsulta(rsx)
           if wtipoLNK = "I" then
              wNomL    = wUrlDesc
              if wNomL <> "" then
                 wNomL = UCASE(wNomL)
              else
                 wNomL = "S/D"
              end if
              if not isnull(wUrl) then
                 sHref = instr(wUrl,"://")
                 if sHref > 0 then
                    sHref = wUrl
                 else
                    sHref = "HTTP://" & wUrl
                 end if
              else
                 sHref = "javascript:void(0);"
              end if
              sFuncion = " target=_blank "
           else
              wNomL    = wUrlDesc
              if wNomL <> "" then
                 wNomL = UCASE(wNomL)
              else
                 wNomL = "S/D"
                 Strsql = "Select * From csGenObjetos where codobjeto = " & wtipoL
                 set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                 if not rsx.eof then
                    wNomL = rsx(1)
                 else
                    wNomL = "S/D"
                 end if
                 call ofv.cerrarconsulta(rsx)

              end if
           end if

           sImagen = wNomL
            SArma = wNomL
            sLink       = ofv.GenLink("","btbn",GenKeyLnks(sHref),"","",GenKeyLnks(sFuncion),sgllnk,"si")
            sLink2       = "<b><i>" & SArma  & "</i></b>" 'ofv.GenLink("","nt",GenKeyLnks(sHref),"","",GenKeyLnks(sFuncion),SArma,"si")
            Celdak = ""
            Celdak = Celdak & ofv.GenCelda("","btbn style='cursor:pointer;'  title='ver documento.'  ","","2","","","",sLink,"si")
            Celdak = Celdak & ofv.GenCelda("","ictx","","","","","",sLink2,"si")
            Filak = ofv.GenRow("","","","","","","",Celdak,"si")
            SArma = ofv.GenTabla("","TWINE","","98%","","0","0","0","-1","","",Filak,"si")
             Celda = ofv.GenCelda("","ictx colspan=3","","20","","","","&nbsp;","si")
             Celda = Celda & ofv.GenCelda("","ictx colspan=4","","","","","",SArma ,"si")
 ''' anula impresion de links ''''''''    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
        end if
       end if
      end if

''' fin links


      if rs2(6) <> "" and rs2(6) > "0"  then
           Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")

           nCopias = "&nbsp;"

           if wstipo(2) then
        'Copias
              if not isnull(rs2(13)) then
                 if rs2(13) > "0" then
                    if rs2(13) > "1" then
                       nCopias = "&nbsp;En&nbsp;" & rs2(13) & "&nbsp;Copias"
                    else
                       nCopias = "&nbsp;En&nbsp;" & rs2(13) & "&nbsp;Copia"
                    end if
                 end if
              end if
           end if

      'Componentes
          sComponente = "&nbsp;"
          sLink       = sglcmp
          sLink2      = "&nbsp;"



          Strsql = "Select nombre From csacomple Where codcomp = " & rs2(6)
          set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if  not rsComponentes.eof then
              sComponente = rsComponentes(0)

                sCompAn = ""
                For oc = 0 To UBound(TAnex)
                    If CLng(TAnex(oc)) = CLng(rs2(6)) Then
                        sCompAn = "&nbsp;(" & TAnexi(oc) & ")&nbsp;"
                        TAnexd(oc) = sComponente
                       Exit For
                    End If
               Next



              SArma       = "<b>" & replace(sComponente," ","&nbsp;")   & sCompAn & nCopias & "</b>"


          end if
          call ofv.cerrarconsulta(rsComponentes)



            Celdak = ""
            Celdak = Celdak & ofv.GenCelda("","ictx","","2","","","","&nbsp;","si")
            Celdak = Celdak & ofv.GenCelda("","ictx","","","","","", SArma   ,"si")

            Filak = ofv.GenRow("","","","","","","",Celdak,"si")
            SArma = ofv.GenTabla("","TWINE","","98%","","0","0","0","-1","","",Filak,"si")

            Celda = Celda & ofv.GenCelda("","ictx colspan=4","","","","","",SArma,"si")
            Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if


      if rs2(8) > "0" or rs2(9) > "0" then
        Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
        'Sector Continuador
        sContinuador = "&nbsp;"
        if rs2(8) > "0" then
          Strsql = "Select descri From Eco Where eco = " & rs2(8)
          set rsContinuadores = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsContinuadores.eof then sContinuador = "Continuador: " & rsContinuadores(0) & "&nbsp;"
          call ofv.cerrarconsulta(rsContinuadores)
        end if


        'Perfil Continuador
        sPerfil = "&nbsp;"
        if rs2(9) > "0" then
          Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs2(9) & "'"
          set rsPerfiles = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsPerfiles.eof then sPerfil = "<b>Perfil: " &rsPerfiles(0) & "&nbsp;<b>"
          call ofv.cerrarconsulta(rsPerfiles)
        end if


        Celdak = ""
        Celdak = Celdak & ofv.GenCelda("","ictx title='sector continuador.' ","","2","","","",sglcnt,"si")
        Celdak = Celdak & ofv.GenCelda("","ictx","","","","","","<b>" & sContinuador & "&nbsp;&nbsp;</b>","si")
        Celdak = Celdak & ofv.GenCelda("","ictx title='perfil continuador.'  ","","2","","","",sglprf,"si")
        Celdak = Celdak & ofv.GenCelda("","ictx","","","","","",sPerfil,"si")


        Filak = ofv.GenRow("","","","","","","",Celdak,"si")
        SArma = ofv.GenTabla("","TWINE","","98%","","0","0","0","-1","","",Filak,"si")

        'Celda = Celda & ofv.GenCelda("","dt ","","","","","","&nbsp;&nbsp;w","si")
        Celda = Celda & ofv.GenCelda("","ictx colspan=4","","","","","",SArma,"si")

        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if
      
      
      
         sLinkalt = "&nbsp;"
            sLink = "&nbsp;"
      ''''      if rs2(5) = "CN" then
       if wstipo(1) then
              'Acciones alternativas
              CantRegAMover = "A"
              
      
               slinkalt = accionesX(nAccionID,nid,nIDd,sTipo,cn)
      
      end if

      if  wstipo(1) then
          Celda =  ofv.GenCelda("","dt colspan=2","","1%","","","","&nbsp;","si")
          Celda =  celda & ofv.GenCelda("","dt colspan=5","","","","","",sLinkalt,"si")
          Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
          sLinkalt = ""
      end if


        Celda    = ""
 
        Celda = Celda & ofv.GenCelda("","dt","","1%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","1%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","3%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","50%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","","","","","&nbsp;","si")

        Fila = Fila & ofv.GenRow("","","","","5","","",Celda,"si")

      rs2.movenext
      posicion = posicion + 1
    loop
    end if


    shtml = shtml & fila   & OFV.GenfTabla()
    
    
    
  fila=""
  celda=""

  rsdd.movenext   
 loop   
  
    sBloque1 = sBloque1 & sHTML   & OFV.GenfTabla()
    
    shtml=""




'-------------------------------------------------
   End If

          ' Cuerpo de Controles
   If Bloque(4) Then
       strsql = "Select Distinct * From Reportes3 Where codobjeto = " & id
       strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
       strsql = strsql & " order by descri, frecuencia,  perdescri, secuencia, descricon "
       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

       hoja = 1
       topeflush = 50

       If Not rs.EOF Then
           Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
           sh = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
           Celda = OFV.GenCelda("", "ibtt", "", "", "", "", "", TBloque(4), "si")
           sh = sh & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
           Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
           sh = sh & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
           sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", sh, "si")

             sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")



           Do Until rs.EOF
               If (rs(0) <> sSector Or rs(1) <> sFrecuencia Or rs(7) <> sPerfil) Then
                          'Encabezado Sector y Frecuencia
                   sSector = rs(0)
                   sPerfil = rs(7)
                   sFrecuencia = rs(1)
             ''        sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")
                   Celda = OFV.GenCelda("", "", "", "1% colspan=5", "", "", "", "&nbsp;", "si")
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                   Celda = OFV.GenCelda("", "istt", "", "70% colspan=3", "", "", "", "Sector - Perfil", "si")
                   Celda = Celda & OFV.GenCelda("", "", "", "1%", "", "", "", "&nbsp;", "si")
                   Celda = Celda & OFV.GenCelda("", "istt", "", "29%", "", "", "", "Frecuencia", "si")
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

                          'Sector
                   sSector = "&nbsp;"
                   If Not IsNull(rs(0)) Then sSector = rs(0)
                   sPerfil = "&nbsp;"
                   If Not IsNull(rs(7)) Then sPerfil = rs(7)
                   Celda = OFV.GenCelda("", "istx", "", "70% colspan=3", "", "", "", sSector & " - " & sPerfil, "si")

                          'Frecuencia
                   sFrecuencia = "&nbsp;"
                   If Not IsNull(rs(1)) Then sFrecuencia = rs(1)
                   Celda = Celda & OFV.GenCelda("", "", "", "1%", "", "", "", "&nbsp;", "si")
                   ssql = "select descri from frecuencia where codfrec='" & Trim(sFrecuencia & "") & "'"
                   Set orsx = OFV.crearconsultaEx(ssql, cn, 1, "")
                          'Celda = Celda & OFV.GenCelda("", "ut", "", "29%", "", "", "", orsx(0), "si")
                   If Not orsx.EOF Then
                       Celda = Celda & OFV.GenCelda("", "istx", "", "90% colspan=4", "", "", "", orsx(0), "si")
                   Else
                       Celda = Celda & OFV.GenCelda("", "istx", "", "90% colspan=4", "", "", "", "DIARIA", "si")
                   End If
                   Call OFV.cerrarconsulta(orsx)
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")




                  'referencia

                   strsql = "Select referencia From objcontrol Where objid = " & rs(2)
                     Set rsx = OFV.crearconsultaEx(strsql, cn, 1, "")

                     If Not rsx.EOF Then
                      sReferencia = "&nbsp;"
                        If Not IsNull(rsx(0)) Then sReferencia = rsx(0)
                        If (Len(Trim(sReferencia)) > 0 And sReferencia <> "&nbsp;") Then


                            Celda = OFV.GenCelda("", "", "", "1% colspan=5", "", "", "", "&nbsp;", "si")
                            sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
  
                            Celda = OFV.GenCelda("", "istt", "center", "90% colspan=5", "", "", "", "Referencia ", "si")
                            sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

                            Celda = OFV.GenCelda("", "ictx", "", "90% colspan=5", "", "", "", "<br>" & Replace(sReferencia & " ", Chr(13), "<br>"), "si")
                            sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

                        End If


                     End If


                   bPrimero = "True"
                   Celda = OFV.GenCelda("", "", "", "1% colspan=5", "", "", "", "&nbsp;", "si")
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
               Else
                   bPrimero = "False"
               End If

                      'Descripcion
               If Not IsNull(rs(5)) Then
                   If bPrimero = "True" Then
                              'Encabezado Descripcion
                       sdesc = 1
                       Celda = OFV.GenCelda("", "istt", "", "1% colspan=5", "", "", "", "Descripcion", "si")
                       sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                       Celda = OFV.GenCelda("", "", "", "1% colspan=5", "", "", "", "&nbsp;", "si")
                       sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                   End If
                   sImage = "&nbsp;"
                   If bVersion = "SI" And rs(8) Then
                       If rs(9) = "M" Then
                           sImage = OFV.GenImage("IC", "", "", "", "Imagenes/Botmodx.jpg", "0", "")
                       ElseIf rs(9) = "A" Then
                           sImage = OFV.GenImage("IC", "", "", "", "Imagenes/Botagrx.jpg", "0", "")
                       End If
                   End If
                   Celda = OFV.GenCelda("", "ictx", "", "1% valign=top", "", "", "", sImage, "si")
                   Celda = Celda & OFV.GenCelda("", "ictx", "", "1% valign=top", "", "", "", rs(6) & ".&nbsp;", "si")
                   Celda = Celda & OFV.GenCelda("", "ictx", "", "99% colspan=3", "", "", "", Replace(rs(5), Chr(13), "<br>"), "si")
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                   Celda = OFV.GenCelda("", "", "", "1% colspan=5", "", "", "", "&nbsp;", "si")
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                   sdesc = sdesc + 1
               End If

               sBloque1 = sBloque1 & sHTML
               sHTML = ""
               rs.movenext
           Loop

             sBloque1 = sBloque1 & OFV.GenfTabla()


       End If
       Call OFV.cerrarconsulta(rs)
   End If

          ' Cuerpo de Esquemas Contables
   If Bloque(5) Then
       strsql = "Select Distinct descri,Eco1descri,evendescri,objid, "
       strsql = strsql & "secuen From Reportes2 Where codobjeto = " & id
       strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
       strsql = strsql & " Order By secuen"
       Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

       hoja = 1
       topeflush = 50
       If Not rs.EOF Then
           Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
           sh = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
           Celda = OFV.GenCelda("", "ibtt", "", "", "", "", "", TBloque(5), "si")
           sh = sh & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
           Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
           sh = sh & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
           sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", sh, "si")

             sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")

           Do Until rs.EOF
               If rs(2) <> sEvento Then
                          'Evento
                   sEvento = rs(2)
              ''       sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")
                   Celda = OFV.GenCelda("", "", "", "1% colspan=2", "", "", "", "&nbsp;", "si")
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                   Celda = OFV.GenCelda("", "istt", "center", "10%", "", "", "", "Evento: ", "si")
                   Celda = Celda & OFV.GenCelda("", "istt", "", "90%", "", "", "", sEvento, "si")
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si") & OFV.GenfTabla()
               End If

                      'Sector Origen
                      'Encabezado Sector Origen y Destino
               sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")
               Celda = OFV.GenCelda("", "", "", "1% colspan=2", "", "", "", "&nbsp;", "si")
               sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
               Celda = OFV.GenCelda("", "", "", "1% colspan=2", "", "", "", "&nbsp;", "si")
               sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
               Celda = OFV.GenCelda("", "istt", "center", "50%", "", "", "", "Sector Origen", "si")
               Celda = Celda & OFV.GenCelda("", "istt", "center", "50%", "", "", "", "Sector Destino", "si")
               sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
               Celda = OFV.GenCelda("", "", "", "1% colspan=2", "", "", "", "&nbsp;", "si")
               sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

               sSector = "&nbsp;"
               If Not IsNull(rs(0)) Then sSector = rs(0)
               Celda = OFV.GenCelda("", "istx", "", "50%", "", "", "", sSector, "si")

               sSector = "&nbsp;"
               If Not IsNull(rs(1)) Then sSector = rs(1)
               Celda = Celda & OFV.GenCelda("", "istx", "", "50%", "", "", "", sSector, "si")
               sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
               Celda = OFV.GenCelda("", "", "", "1% colspan=2", "", "", "", "&nbsp;", "si")
               sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                      ' Ver referencia del Esquema Contable
               sReferencia = ""
               strsql = "SELECT OBJESQUEMA.refesq FROM OBJESQUEMA where objid = " & rs(3)
               Set rsCuenta = OFV.crearconsultaEx(strsql, cn, 1, "")
               If Not rsCuenta.EOF Then
                   sReferencia = rsCuenta(0)
                   If Len(sReferencia) > 0 Then
                       Celda = OFV.GenCelda("", "istt", "center COLSPAN=2", "", "", "", "", "Observaciones", "si")
                       sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                       Celda = OFV.GenCelda("", "istx", "", "100% COLSPAN=2", "", "", "", sReferencia, "si")
                       sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                   End If
               End If

               sHTML = sHTML & OFV.GenfTabla()

               strsql = "SELECT OBJCUENTA.descri, OBJCUENTA.numero, OBJESQUEMA2.sentido, "
               strsql = strsql & "OBJESQUEMA2.refesq2, OBJESQUEMA2.modi, OBJESQUEMA2.tipmodi "
               strsql = strsql & "FROM OBJESQUEMA2, OBJCUENTA where "
               strsql = strsql & "OBJESQUEMA2.codcuenta = OBJCUENTA.objid "
               strsql = strsql & "and Objesquema2.codesquema = " & rs(3)
               strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
               Set rsCuenta = OFV.crearconsultaEx(strsql, cn, 1, "")
               If Not rsCuenta.EOF Then
                          'Encabezado Cuentas
                   sHTML = sHTML & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")
                   Celda = OFV.GenCelda("", "istt colspan=2", "", "", "", "", "", "Sentido", "si")
                   Celda = Celda & OFV.GenCelda("", "istt", "", "", "", "", "", "Cuentas Contables", "si")
                   Celda = Celda & OFV.GenCelda("", "istt", "", "", "", "", "", "Referencia", "si")
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                   Celda = OFV.GenCelda("", "", "", "1% colspan=4", "", "", "", "&nbsp;", "si")
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

                   Do Until rsCuenta.EOF
                       sImage = "&nbsp;"
                       If bVersion = "SI" And rsCuenta(4) Then
                           If rsCuenta(5) = "M" Then
                               sImage = OFV.GenImage("IC", "", "", "", "Imagenes/Botmodx.jpg", "0", "")
                           ElseIf rsCuenta(5) = "A" Then
                               sImage = OFV.GenImage("IC", "", "", "", "Imagenes/Botagrx.jpg", "0", "")
                           End If
                       End If
                       Celda = OFV.GenCelda("", "ictx", "", "3% valign=top", "", "", "", sImage, "si")
                       Celda = Celda & OFV.GenCelda("", "icxr", "", "8%", "", "", "", rsCuenta(2), "si")
                       Celda = Celda & OFV.GenCelda("", "ictx", "", "40%", "", "", "", rsCuenta(1) & " - " & rsCuenta(0), "si")
                       sReferencia = rsCuenta(3)
                       If IsNull(sReferencia) Then sReferencia = "&nbsp;"
                       Celda = Celda & OFV.GenCelda("", "ictx", "", "49%", "", "", "", sReferencia, "si")
                       sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                       rsCuenta.movenext
                   Loop
                   Celda = OFV.GenCelda("", "", "", "1% colspan=4", "", "", "", "&nbsp;", "si")
                   sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
               End If
               Call OFV.cerrarconsulta(rsCuenta)

               sBloque1 = sBloque1 & sHTML
               sHTML = ""

               rs.movenext
           Loop

             sBloque1 = sBloque1 & OFV.GenfTabla()

       End If
       Call OFV.cerrarconsulta(rs)
   End If


          ' Cuerpo de Tablas
   If Bloque(6) Then

       Nblanco = OFV.GenCelda("", "", "", "100%", "", "", "", "&nbsp;", "si")
       Nblanco = OFV.GenRow("", "", "", "", "", "", "", Nblanco, "si")
       Nblanco = OFV.GenTabla("TB", "", "", "600", "", "0", "0", "0", "0", "", "", Nblanco, "si")

       strsql = "SELECT * From Objtabdef Where codobjeto = " & id & " order by tabsecuen"
       
       Set rs1 = OFV.crearconsultaEx(strsql, cn, 1, "")

       Do Until rs1.EOF
           nTabla = rs1(0)
           nTitulo = rs1(4)
           nPerfil = rs1(3)

                  'Bloque de Defaults
           ncolcant = 1
           nBorde = False
           nblkfc = Session("CMIETTFT")
           nblkbc = Session("CMIETT")
           nblkff = Session("CMIETTFTFC")
           nblkfs = Session("CMIETTFTSZ") & "pt"

           nctitfc = Session("CMISTTFT")
           nctitbc = Session("CMISTT")
           nctitff = Session("CMISTTFTFC")
           nctitfs = Session("CMISTTFTSZ") & "pt"

           ncconfc = Session("CMICTXFT")
           ncconbc = Session("CMICTX")
           ncconff = Session("CMICTXFTFC")
           ncconfs = Session("CMICTXFTSZ") & "pt"

           nStilo = " Style={font-family:" & nblkff & ";font-size:" & nblkfs & ";color:#993333;background-color:#ffffff;font-weight:bold;} "

           nfontx = "</font>"

                  '  nStilo = " Style={font-family:" & nblkff & ";font-size:" & nblkfs & ";color:#993333;background-color:#ffffff;font-weight:bold;} "

           sHTML = sHTML & Nblanco
           Celda = OFV.GenCelda("", "ibtt", "", "95%", "", "", "", "Tabla:&nbsp;" & nTitulo, "si")
           Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
           sHTML = sHTML & OFV.GenTabla("TB", "", "", "600", "", "0", "", "0", "0", "#ffffff", "", Fila, "si")
           sHTML = sHTML & Nblanco



           ReDim aColumnas(0)
           ReDim aColumnasa(0)

           strsql = "Select * From Tabperfil Where Tabperfil = '" & nPerfil & "'"
          
           Set rs2 = OFV.crearconsultaEx(strsql, cn, 1, "")
           If Not rs2.EOF Then
               ncolcant = CLng(rs2(7))

               nBorde = rs2(2)
               nblkfc = rs2(3)
               nblkbc = rs2(4)
               nblkff = Replace(rs2(5), " ", "-")
                      'nblkfs = rs2(6)
               nblkfs = cnumer(CStr(rs2(6))) & "pt"

               nctitfc = rs2(8)
               nctitbc = rs2(9)
               nctitff = Replace(rs2(10), " ", "-")
                      'nctitfs = rs2(11)
               nctitfs = cnumer(CStr(rs2(11))) & "pt"


               ncconfc = rs2(12)
               ncconbc = rs2(13)
               ncconff = Replace(rs2(14), " ", "-")
                      'ncconfs = rs2(15)
               ncconfs = cnumer(CStr(rs2(15))) & "pt"

               ReDim Preserve aColumnas(ncolcant - 1)
               ReDim Preserve aColumnasa(ncolcant - 1)
               
               
                redim preserve aColumnascol(cint(ncolcant) - 1)
                redim preserve aColumnasrow(cint(ncolcant) - 1)
  
  
                for ix = 0 to UBound(aColumnas)
	               aColumnas(ix) = rs2(ix + 16)
	               aColumnasa(ix) = rs2(ix + 26)
	                 aColumnascol (ix) = 0
	                 aColumnasrow (ix) = 0
                 next
           End If

           Call OFV.cerrarconsulta(rs2)
           ib=1
	    redim preserve aBloques(0)
	    redim preserve aBloquess(0)
	    redim preserve aBloquest(0)
            xhtml = ""
            'aca secuencia titulo
           strsql = "SELECT * From Objtabblk Where Tabdef = " & nTabla & " order by tabblksec"
        
           Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
           Do Until rs.EOF
               nid = rs(0)
               nSec = rs(4)
               nBTitulo = rs(5)
                 wcolspan  = "N" 
	         wrowspan  = "N" 
	         wcolspanc = 0
	         wrowspanc = 0
	         wcolspann = 0
                 wcolspanw = "N"
  
               If Len(nBTitulo) > 0 Then
                   ltitu = Len(nBTitulo)
                   If nBTitulo <> Space(ltitu) And UCase(nBTitulo) <> "NO" Then
                      redim preserve aBloques(ib)
		      redim preserve aBloquess(ib)
		      redim preserve aBloquest(ib)

		      aBloques(ib)  = rs(0)
		      aBloquess(ib) = rs(4)
		      aBloquest(ib) = rs(5)

		      sPoint  = "BLK" & aBloques(ib)

                       ib = ib + 1
           

                       nStilo = " Style={font-family:" & Chr(34) & nblkff & Chr(34) & ";font-size:" & nblkfs & ";color:" & nblkfc & ";background-color:" & nblkbc & ";text-decoration:none;} "

                       nfonty = " <font face=" & Chr(34) & Replace(nblkff, "-", " ") & Chr(34) & " color=" & nblkfc & " >"

                              '  nStilo = " Style={font-family:" & nblkff & ";font-size:" & nblkfs & ";color:" & nblkfc & ";background-color:" & nblkbc & ";} "

                       sHTML = sHTML & Nblanco
                       Celda = OFV.GenCelda("", "aaa" & nStilo, "", "5%", "", "", "", nfonty & nSec & nfontx, "si")
                       Celda = Celda & OFV.GenCelda("", "aaa" & nStilo, "", "95%", "", "", "", nfonty & nBTitulo & nfontx, "si")
                       Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                       sHTML = sHTML & OFV.GenTabla("TB", "", "", "600", "", "0", "0", "0", "0", "#ffffff", "", Fila, "si")
                       sHTML = sHTML & Nblanco
                   End If
               End If
               Fila = ""

               strsql = "SELECT * From Objtabcon Where tabblk = " & nid & " order by tabcontipo,tabconsec"
               Set rs2 = OFV.crearconsultaEx(strsql, cn, 1, "")

               Do Until rs2.EOF

                   nStilot = " Style={font-family:" & Chr(34) & nctitff & Chr(34) & ";font-size:" & nctitfs & ";color:" & nctitfc & ";background-color:" & nctitbc & ";} "
                   nfontt = " <font face=" & Chr(34) & Replace(nctitff, "-", " ") & Chr(34) & " color=" & nctitfc & " >"


                   nStiloc = " Style={font-family:" & Chr(34) & ncconff & Chr(34) & ";font-size:" & ncconfs & ";color:" & ncconfc & ";background-color:" & ncconbc & ";} "
                   nfontc = " <font face=" & Chr(34) & Replace(ncconff, "-", " ") & Chr(34) & " color=" & ncconfc & " >"

                          '  nStilot = " Style={font-family:" & nctitff & ";font-size:" & nctitfs & ";color:" & nctitfc & ";background-color:" & nctitbc & ";} "
                          '  nStiloc = " Style={font-family:" & ncconff & ";font-size:" & ncconfs & ";color:" & ncconfc & ";background-color:" & ncconbc & ";} "


                
                 for ix = 0 to UBound(aColumnas)
		 
		           aColumnascol (ix) = 0
		           aColumnasrow (ix) = 0
		 
                  next




''''''''''''''''''''''''''
    strsql = "SELECT * From Objtabcmb Where TABobj = '" & rs2(26) & "'  "
    strsql = strsql & " order by tabcol "   
 ''   response.write strsql & "<br>" & nid & "<br>"
    set rs21 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    do until rs21.eof 

       if  rs21(8) <> "N" then aColumnasrow ((clng(rs21(6)) - 1)) = rs21(8)

       if  rs21(8) <> "N" then
           if  wcolspanw  = "N" then
               if  wrowspan  = "N" then
                   wrowspanc = rs21(8) 
                   wcolspanc = rs21(7) 
                   wcolspann = clng(rs21(6))
                   wcolspanw = "R"

                   if wcolspanc = "N" then wcolspanc = 1
                   if wrowspanc = "N" then wrowspanc = 1


               end if
           end if
       end if

       if  rs21(7) <> "N" then aColumnascol ((clng(rs21(6)) - 1)) = rs21(7)
       if  rs21(7) <> "N" then
           if  wcolspanw  = "N" then
               if  wcolspan  = "N" then
                   wcolspanc = rs21(7) 
                   wcolspann = clng(rs21(6))
                   wcolspanw = "S"

                   if wcolspanc = "N" then wcolspanc = 1

               end if
           end if

       end if





       rs21.movenext
    loop 
''''''''''''''''''''''''''





                
                   nStilo = nStiloc
                   nFontr = ncconbc
                   nfonta = nfontc
                   If rs2(5) = "1" Then
                       nStilo = nStilot
                       nFontr = nctitbc
                       nfonta = nfontt
                   End If





   nStilo = nStilo & cstil &  "} "
   
         wcsc = 0


  'Columnas
      celda = ""
      for ix = 0 to UBound(aColumnas)


''''''''''''''''''

        wxrs = ""
        wxcs = ""

        if  wcolspanw = "R" then
            if  wrowspan  = "N" and ix = (wcolspann-1) then
                wrowspan  = "S" 
                wcolspan  = "S"

                wxrs = wrowspanc
                wxcs = wcolspanc

                if wxrs = 0 then wxrs = 1
                if wxcs = 0 then wxcs = 1

                wxrs = " ROWSPAN=" & wxrs & " "
                wxcs = " COLSPAN=" & wxcs & " "

                wcsc = 101



            elseif  wrowspan  = "S" and ix = (wcolspann-1) then

                wxrs = ""

                if wcolspanc > 1 then 
                   wcsc = wcolspanc
                else
                   wcsc = 1
                end if

            end if 




        end if 

        if  wcolspanw = "S"  then
            if  wcolspan  = "N" and ix = (wcolspann-1) then
                wcolspan  = "S" 

                
                wxrs = ""
                wxcs = wcolspanc

                if wxcs = 0 then wxcs = 1

                wxcs = " COLSPAN=" & wxcs & " "

                wcsc = 101



            end if 
        end if 

        if  wcsc = 0 or wcsc > 100 then


''''''+++++++++++++++++++++++

       '     sInput = rs2(ix + 6)
       '     Celda = Celda & ofv.GenCelda("","vt "  & cstil & wxrs & " " & wxcs ,"","","","","",sinput ,"si")


           sSize = cstr(aColumnas(ix)) & "% "
           sConten = rs2(ix + 6)
           if isnull(sconten) then sconten = " "
           lconten = len(sconten) 
           if lconten > 0 then
              if sConten = "" or sConten = space(lconten) then sConten = "&nbsp;"
           end if   
           if lconten = 0 then sConten = "<font color=" & nFontr & ">-</font>" 

           sConten = nfonta &  replace(sConten,chr(13),"<br>") & nfontx
           select case cstr(aColumnasa(ix))
                  case "0" 
                         xalig = " align=center "
                  case "1" 
                         xalig = " align=left "
                  case "2" 
                         xalig = " align=right "
           end select

           Celda = Celda & ofv.GenCelda("","aaa" & nStilo & xalig & " valign=top " & wxrs & " " & wxcs,"",sSize,"","","",sConten,"si")




'''''''++++++++++++++++++++++++++


            if  wcsc > 100 then
                if wcolspanc > 1 then 
                   wcsc = wcolspanc - 1
                else
                   wcsc = 0
                end if
            end if

        else

            wcsc = wcsc - 1

            if  wcsc = 0 then
                wcolspan  = "N"
            end if 

            if  wcolspanw = "S" and wcolspan  = "N" then

                wcolspanw = "N"

                for ix2 = ix + 1 to UBound(aColumnas)

                    if aColumnascol (ix2) <> 0 then
                       wcolspanc = aColumnascol (ix2)
                       wcolspann = ix2 + 1
                       wcolspanw = "S"
                       exit for
                    end if 

                next

            end if 

        end if





'''''''''''''''''''

      next
      
      
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
      
            if  wrowspan  <> "N" then
      
                wrowspanc = wrowspanc - 1
      
                if  wrowspanc = 0 then
                    wrowspan  = "N"
                    wcolspann = 0
                    wcolspanw = "N"
                    wcolspanc = 0 
                end if
      
      
            end if 
      

                   rs2.movenext
               Loop
               If nBorde Then
                   sHTML = sHTML & OFV.GenTabla("TB", "", "", "600 border", "", "0", "0", "0", "0", "#ffffff", "", Fila, "si")
               Else
                   sHTML =  sHTML & OFV.GenTabla("TB", "", "", "600", "", "0", "0", "0", "0", "#ffffff", "", Fila, "si")
               End If
               sHTML = sHTML & Nblanco

               Call OFV.cerrarconsulta(rs2)

               rs.movenext
           Loop

           Call OFV.cerrarconsulta(rs)
           rs1.movenext
       Loop
       sBloque1 = sBloque1 & sHTML 
       sHTML = ""
       Call OFV.cerrarconsulta(rs1)
   End If

          ' Anexos de Complementos
   If Bloque(3) Or Bloque(7) Then
       For oc = 0 To UBound(TAnex)
           If TAnexc(oc) = "s" Then
               For z = 0 To UBound(TAnexa)
                   If CLng(Tanexe(z)) > CLng(TAnex(oc)) Then Exit For
                   If CLng(Tanexe(z)) = CLng(TAnex(oc)) Then
                       sHTML = sHTML & "<div>&nbsp;</div>" & OFV.GenTabla("", "", "", "600", "", "0", "0", "0", "0", "", "", "", "no")
                       Celda = OFV.GenCelda("", "", "", "1% colspan=2", "", "", "", "&nbsp;", "si")
                       sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                       Celda = OFV.GenCelda("", "istt", "center", "10%", "", "", "", "&nbsp;" & TAnexi(oc) & ":&nbsp;", "si")
                       Celda = Celda & OFV.GenCelda("", "istt", "", "90%", "", "", "", TAnexd(oc) & TAnexf(z) & "&nbsp;", "si")
                       sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
                       Celda = OFV.GenCelda("", "", "", "1% colspan=2", "", "", "", "&nbsp;", "si")
                       sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si") & OFV.GenfTabla()
                       sHTML = sHTML & OFV.GenImage("Imagen", "", "", "", "Componentes/" & TAnexa(z), "1", "")
                   End If
               Next
           End If
       Next
       sBloque1 = sBloque1 & sHTML
       sHTML = ""

   End If


   Call OFV.cerrarconn(cn)


   Generar = sBloque1
     

        End Function


Public Function ConvNumer(CNumero, CTipo)

  

    Dim CLetra, ci, cx, cl, cr
    
    CLetra = "abcdefghijklmnopqrstuvwxyz"

    Dim CRuni(9)
    CRuni(1) = "I"
    CRuni(2) = "II"
    CRuni(3) = "III"
    CRuni(4) = "IV"
    CRuni(5) = "V"
    CRuni(6) = "VI"
    CRuni(7) = "VII"
    CRuni(8) = "VIII"
    CRuni(9) = "IX"

    Dim CRdec(9)
    CRdec(1) = "X"
    CRdec(2) = "XX"
    CRdec(3) = "XXX"
    CRdec(4) = "XL"
    CRdec(5) = "L"
    CRdec(6) = "LX"
    CRdec(7) = "LXX"
    CRdec(8) = "LXXX"
    CRdec(9) = "XC"

    Dim CRcen(9)
    CRcen(1) = "C"
    CRcen(2) = "CC"
    CRcen(3) = "CCC"
    CRcen(4) = "CD"
    CRcen(5) = "D"
    CRcen(6) = "DC"
    CRcen(7) = "DCC"
    CRcen(8) = "DCCC"
    CRcen(9) = "CM"


    Select Case CTipo
        Case "a", "A"
            ci = 0
            cr = CNumero

            For ci = 0 To 28
                If cr < 28 Then Exit For
                cr = cr - 27
            Next

            If ci > 27 Then
                ConvNumer = "XX"
            Else
                If ci > 0 Then ConvNumer = Mid(CLetra, ci, 1)
                ConvNumer = ConvNumer & Mid(CLetra, cr, 1)
            End If
            If CTipo = "A" Then ConvNumer = UCase(ConvNumer)

        Case "I", "i"
            cr = CNumero
            ci = Len(cr)
            If ci < 1 Or ci > 3 Then
                ConvNumer = "AAAA"
            Else
                cx = 1
                For cl = ci To 1 Step -1
                    Select Case cx
                        Case 1
                            ConvNumer = CRuni(Mid(cr, cl, 1))
                        Case 2
                            ConvNumer = CRdec(Mid(cr, cl, 1)) & ConvNumer
                        Case 3
                            ConvNumer = CRcen(Mid(cr, cl, 1)) & ConvNumer
                    End Select

                    cx = cx + 1
                Next
                If CTipo = "i" Then ConvNumer = LCase(ConvNumer)
            End If
    End Select

   

End Function

Public Function RepHeader2(STitulo, nAncho) 

'''        Public RCliLogo 
'''        Public RCliNomb 
    Dim s, Fecha
    
    If nAncho = "" Then nAncho = "600"
  
    s = "<table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
    s = s & "<tr align=top>"
    If RCliLogo <> "" Then s = s & "<td width=4% ><img src=" & RCliLogo & " border=0 ></td>"
    s = s & "<td class=ietx align=left >&nbsp;" & RCliNomb & "</td></tr></table>"
    s = s & "<br><table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
    s = s & "<tr align=top><td class=iett align=left width=70% >" & STitulo & "</td>"
    Fecha = OFV.fechador()
    s = s & "<td class=ietx align=right>" & Fecha & "</td></tr></table>"
    s = s & "<hr color=#000088 size=1 width=" & nAncho & " align=left>"

    RepHeader2 = s

End Function
Function cnumer(cdato ) 
    Dim ldato 
    Dim kk 
   
    ldato = Len(cdato)

    For kk = 1 To ldato
        If IsNumeric(Mid(cdato, kk, 1)) Then
            cnumer = cnumer & Mid(cdato, kk, 1)
        End If
    Next


End Function
Function CountOccurrences(strText ,strFind , lngCompare ) 

    ' Count occurrences of a particular character or characters.
    ' If lngCompare argument is omitted, procedure performs binary comparison.
   
    Dim lngPos       
    Dim lngTemp      
    Dim lngCount     
   
    ' Specify a starting position. We don't need it the first
    ' time through the loop, but we'll need it on subsequent passes.
    lngPos = 1
    ' Execute the loop at least once.
    Do
        ' Store position at which strFind first occurs.
        lngPos = InStr(lngPos, strText, strFind, lngCompare)
        ' Store position in a temporary variable.
        lngTemp = lngPos
        ' Check that strFind has been found.
        If lngPos > 0 Then
            ' Increment counter variable.
            lngCount = lngCount + 1
            ' Define a new starting position.
            lngPos = lngPos + Len(strFind)
        End If
        ' Loop until last occurrence has been found.
    Loop Until lngPos = 0
    ' Return the number of occurrences found.
    CountOccurrences = lngCount
End Function

'Function getReadfilez(ROOT, url)
'  Dim xmlHttp 
'  Dim ss 
'  Set xmlHttp = Server.CreateObject("MSXML2.XMLHTTP")
'   url = Session("UrlGetdata") & "?OP=F&xarchivo=" & ROOT
'   '    Response.write url
'   xmlHttp.Open "Get", url, False, "", ""
'   Call xmlHttp.Send
'   ss = xmlHttp.ResponseText
'   getReadfile = ss '& session("sitecor")
'  
  
'End Function



function accionesX(nAccion,nCodigo,nProced,sTipo,cn)
  dim shtmlXX
  dim rscc,rs2
  'set cn = ofv.conectar(ofv.strconn2)

  strsqlx = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rscc = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  if not rscc.eof then
    if rscc(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rscc)


  strsql = "SELECT TIPOACCALT.tipoaccionalt, TIPOACCALT.descri, Objaccion.secuen "
  strsql = strsql & "FROM OBJACCION, TIPOACCALT where "
  strsql = strsql & "OBJACCION.tipoaccion = TIPOACCALT.tipoaccionalt "
  strsql = strsql & "and OBJACCION.codaccion = '" & nAccion & "'"
  strsql = strsql & " Order By secuen"
  'response.write strsql
  set rscc = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rscc.eof then
     VTACC = rscc(0)
     VDACC = rscc(1)
     VSACC = rscc(2)
  end if
  call ofv.cerrarconsulta(rscc)




  strsql = "SELECT * from csaccaltsel "
  strsql = strsql & "Where idaccion = '" & nAccion & "' "
  strsql = strsql & " and (modi = 0 or "
  strsql = strsql & " (modi <> 0 and tipmodi <> 'E')) "
  strsql = strsql & " order by secuencia"
  'response.write strsql
  set rscc = ofv.crearconsultaEx(StrSql,cn,1,volver)



  strsql = "SELECT * FROM TIPOACCALT WHERE TIPOACCALT.modo='OP' Or TIPOACCALT.modo='RF' "
  strsql = strsql & "ORDER BY descri"


  sHTMLxx = ""

  'sHTMLxx = sHTMLxx & ofv.GenTabla("panel","aaa style='border-style:solid;border-width:1;border-color:#eeeeee;' ","","80% ","","0","0","-1","2","","","","no")
  sHTMLxx = sHTMLxx & ofv.GenTabla("panel","aaa style='border-style:solid;border-width:1;border-color:#eeeeee;' ","","80% ","","0","0","-1","2","","","","no")

   t2    = ofv.GenCelda("","utis colspan=3 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' ","","","","","",sfidkC & "<b>&nbsp;NO.</b>","si")
  ' t2    = ofv.GenCelda("","utis colspan=3  ","","","","","",sfidkC & "<font color=#ff0000 ><b>&nbsp;NO.</b></font>","si")
 
  sHTMLxx = sHTMLxx & ofv.GenRow("","","","","","","",t2,"si")


  Vcont = 1
  Valor = ""
  VCTACC = ""
  VCDACC = ""
  VCSACC = ""

  if not rscc.eof then
  do until rscc.eof
    Valor = rscc(0)
    if not isnull(rscc(8)) then
       Continuador = rscc(8)


        strsql = "SELECT TIPOACCALT.tipoaccionalt, TIPOACCALT.descri, Objaccion.secuen "
        strsql = strsql & "FROM OBJACCION, TIPOACCALT where "
        strsql = strsql & "OBJACCION.tipoaccion = TIPOACCALT.tipoaccionalt "
        strsql = strsql & "and OBJACCION.codaccion = '" & Continuador & "'"
        strsql = strsql & " Order By secuen"
       'response.write strsql
       set rs2x = ofv.crearconsultaEx(StrSql,cn,1,volver)
       if not rs2x.eof then
          VCTACC = rs2x(0)
          VCDACC = rs2x(1)
          VCSACC = rs2x(2)
       end if
       call ofv.cerrarconsulta(rs2x)

    end if
    ID = rscc(2)
      sImage = "&nbsp;"
      if bVersion = "SI" and rscc(9) then
      if rscc(10) = "M" then
        sImage = ofv.GenImage("","","","","Imagenes/Botmodx.jpg","0","")
      elseif rscc(10) = "A" then
        sImage = ofv.GenImage("","","","","Imagenes/Botagrx.jpg","0","")
      end if
      end if

    t1 = ofv.GenCelda("","dt","","5%","","","",sImage,"si")

    t1 = t1 & ofv.GenCelda("","dt","","5%","","","","&nbsp;" & VSACC & "." & rscc(6) & "&nbsp;","si")

    Vcont = Vcont + 1

     'Texto Accion Alt
      if rscc(4) <> "" then
        Strsql = "Select descri From Tipoaccalt Where tipoaccionalt = " & rscc(4)
        'response.write strsql
        set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsAcciones.eof then
           sAccion = rsAcciones(0)
        end if
        call ofv.cerrarconsulta(rsAcciones)
      end if

      sVerbo = sAccion
      sAcciont = sAccion
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo =  "<b>" & left(sAccion,i-1)  & "</b>"
        sAcciont = mid(sAccion,i+1)
      end if
      sInput = "<i>" & sVerbo  & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & sAcciont 

    t2 = ofv.GenCelda("","ictx ","","90%","","","",sInput,"si")

    sHTMLxx = sHTMLxx & ofv.GenRow("","","","","","","",t1 & t2,"si")

    'Texto de Referencia
    if ofv.explorador = "MSIE" then
      nCol = 60
    else
      nCol = 60
    end if

    t3 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t3 = t3 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    sArma = rscc(7) & chr(13)
    sarma= sarma 
    if len(sArma) > 0 then
       sArma = replace(sArma,chr(13),"<br>")
    end if
'''    SArma = "<font color=#225588 face=arial-narrow size=1>" & sArma & "</font>"
    t3 = t3 & ofv.GenCelda("","ictx ","","","","","",SArma,"si")
    sHTMLxx = sHTMLxx & ofv.GenRow("","","","","","","",t3,"si")
    rscc.movenext
  loop


    'Encabezado Continuador



    t4 = ofv.GenCelda("","","","2%","","","","&nbsp;","si")
    t4 = t4 & ofv.GenCelda("","","","","","","","&nbsp;","si")
      'Texto Continuador
    if VCDACC <> "" then
      sVerbo = VCDACC
      sAcciont = VCDACC
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo =  "<b>" & left(VCDACC,i-1)  &"</b>"
        sAcciont = mid(VCDACC,i+1)
      end if
      sInput = "<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & sAcciont
      sContinuador = "<b>" & VCSACC & " - " & sInput & "</b>"
    end if
    t4 = t4 & ofv.GenCelda("","et","","","","","","Continua en: " & sContinuador,"si")
    if VCDACC <> "" then
       sHTMLxx = sHTMLxx & ofv.GenRow("","","","","","","",t4,"si")
    end if

else

    t6 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t6 = t6 & ofv.GenCelda("","","","","","","","&nbsp;","si")
 ''   SArma = "<h5><B>Sin&nbsp;datos&nbsp;disponibles</B></h5>"
    SArma = "&nbsp;"
    t6 = t6 & ofv.GenCelda("","dt","","","","","",SArma,"si")
    sHTMLxx = sHTMLxx & ofv.GenRow("","","","","","","",t6,"si")


end if

    t5 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    sHTMLxx = sHTMLxx & ofv.GenRow("","","","","","","",t5,"si")


     sHTMLxx = shtmlxx & ofv.GenFTabla()

  call ofv.cerrarconsulta(rscc)



  accionesX=shtmlxx
end function


end class
%>