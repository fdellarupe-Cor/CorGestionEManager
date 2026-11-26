
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok = OFV.CheckUsuario()
If ok <> "Y" Then
    uv = OFV.uv
    If uv = "NO" Then ok = OFV.ValidarUsuario
    shtml = ok
Else
    response.expires = 0
    'response.write session("codSol")
    OFV.ObtenerAtributos session("Form2"), ""
    bModificar = OFV.AModificar
    If session("Nomodificar") = "S" Then bModificar = False


    shtml = OFV.MenuHeader("", "")

    spuesto = request.querystring("puesto")

    xTrColor = "#FFFFFF"

    sImageva = ""   '"<font face='corflaticon' size=2 color='#00aa00'><b>&#0252;</b></font>" & "&nbsp;&nbsp;"
    sImageIT = "<font face='corflaticon' size=2 color='#FF0000'><b>&#0167;</b></font>" & "&nbsp;&nbsp;"

    wkstylof = "aaa style='font-family:arial;font-size:6pt;color:#888888;' "

    BCOLORG = "#cccccc"
    BCOLORD = "#6699CC"

    '''''''''''''''''''''''''''''

    wkstylo1 = " style='background-color:#ddddff;color:#000000;font-weight:bold;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "
    wkstylo2 = " style='background-color:#f7eedd;color:#000000;font-weight:bold;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "
    wkstylo3 = " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' "


    Set cn = OFV.conectar(OFV.strconn2)


    'response.write "<center>"

    Fila = ""

    CeldaD = ""
    CeldaD = CeldaD & OFV.GenCelda("", "utI  ", "", "95%", "", "", "", "DESCRIPCION DE FUNCIONES DEL PUESTO", "si")
    FilaD = OFV.genrow("", "", "", "", "", "", "", CeldaD, "si")
    sFORMhed = OFV.GenTabla("", "TWIND", "", "100%", "", "0", "", "0", "1", "", "", FilaD, "si")


    '' shtml = shtml & sFORMhed

    shtml = shtml & "<BR>"


    Fila = ""





    StrSql = "Select Ecoperfil.*,listadepuestos.descri,ISNULL(E.DESCRI,'SIN DEFINIR') as ecodescri,listadepuestos.NIVELJ,listadepuestos.tipo,listadepuestos.NIVELJreal From Ecoperfil, listadepuestos "
    StrSql = StrSql & " LEFT JOIN ECO E ON(E.ECO=listadepuestos.SECTOR) "
    '  STRSQL= STRSQL & " LEFT JOIN RELAPUESTO EX ON(EX.PUESTODEPEN=listadepuestos.OBJID)  "
    StrSql = StrSql & " where "
    StrSql = StrSql & "Ecoperfil.codsector = listadepuestos.objid and Ecoperfil.codsector= '" & spuesto & "'  and Ecoperfil.codperfil = '****PUESTO' "
   
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, parametros)



    If rs.EOF Then
        ecodescri = "SIN DEFINIR"

        Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
        Fila = Fila & OFV.genrow("", "", "", "", "20", "", "", Celda, "si")

        Celda = OFV.GenCelda("", "uti colspan=2 style='border-style:solid;border-color:#888888;border-width:1;'  ", "center", "", "", "", "", "<big>No hay funciones descriptas para el puesto.</big>", "si")
        Fila = Fila & OFV.genrow("", "", "", "", "50", "", "", Celda, "si")

    Else

        ecodescri = rs("ecodescri")
        sperfil = rs(0)
        NIVELJ = rs("NIVELJreal")
        tipo = rs("tipo")
        Celda = OFV.GenCelda("", "UTI style='color:#ffffff;' ", "", "30%", "", "", "", "DENOMINACION DEL PUESTO:", "si")


        SSECTOR = rs(1)
        SSECTORDAT = rs(1)

        xgrafica = "N"
        spsector = ""
        StrSql = "Select * From LISTADEPUESTOS where OBJID = 0" & rs(1)
        Set rsX = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
        If Not rsX.EOF Then

            '   Strsql = "Select grafica From nivelPUESTO where NIVELP = 0" & NIVELJ
            '  set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            '    if Not rsX2.EOF then xgrafica = rsX2(0) & ""
            xgrafica = rsX("grafica") & ""
            SSECTOR = "<font face=verdana color=#FFFFFF size=3><b><U>" & rsX(1) & "</U></b></font>"    'RSX(1)
            spsector = rsX(1)


            'if     xgrafica = "S" then
            If tipo = "XG" Then
                shref = "OrganigramaPuestosgraf.asp?ID=" & rs(1) & "&Sector=" & rsX(1)
                sinput = "<font face=verdana color=#FFFFFF size=3><b><U>" & SSECTOR & "</U></b></font>"
                slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

                SSECTOR = slink

            End If





        End If

        'SECTOR
        sinput = SSECTOR
        Celda = Celda & OFV.GenCelda("", "utI style='color:#ffffff;' ALIGN='RIGHT' ", "", "70%", "", "", "", sinput & "&nbsp;&nbsp;&nbsp;&nbsp;", "si")

        Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")

    End If


    shtml = shtml & "<center>" & OFV.GenTabla("", "AAA style='border-style:solid;border-color:#cccccc;border-width:1;background-color:#888888;' ", "", "95%   ", "", "0", "", "1", "1", "", "", Fila, "si") & "</center>"
     
     
      BTIT1 = "<A CLASS=CC HREF='BSQINFO.ASP?BSQ=PORPUESTO&ID="&spuesto&"'>Ver Datos de los Ocupantes del Puesto</A>"
      BTIT2 = "<A CLASS=CC HREF='BSQINFO.ASP?BSQ='>Ver el Régimen de Firmas Asociado al Puesto</A>"
             
     Celda = OFV.GenCelda("", "UTI   " & wkstylo1, "", "", "", "", "", BTIT1, "si")
     Celda = CELDA & OFV.GenCelda("", "UTI   " & wkstylo1, "", "", "", "", "", BTIT2, "si")
     
     Fila = OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")
     
    shtml = shtml & "<BR><center>" & OFV.GenTabla("", "AAA style='border-style:solid;border-color:#cccccc;border-width:1;background-color:#888888;' ", "", "", "", "0", "", "1", "1", "", "", Fila, "si") & "</center>"




    Formulario = ""


    '''''''' PERTENECE A:


    '   Formulario = ""


    Fila = ""

    posi = 0
    '  IF Not rs.EOF THEN

    Celda = OFV.GenCelda("", "UTI   " & wkstylo1, "", "20%", "", "", "", "PERTENECE A:", "si")




    SQL = "SELECT a.PUESTOMASTER,e.descri,e.grafica FROM RELAPUESTO a "
    SQL = SQL & " inner join  listadepuestos e on (e.objid=a.PUESTOMASTER) WHERE  a.PUESTODEPEN=0" & spuesto & " ORDER BY grafica "
    'response.write sql &"<br>"
    Set rsx2PX = OFV.crearconsultaEx(SQL, cn, 1, parametros)
    If Not rsx2PX.EOF  Then
        PUESTOMASTER = rsx2PX(0)
        PUESTOMASTERdescri = rsx2PX(1)
        'xgrafica=rsx2PX(2)
       ' Set rsX = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
	
	    Do While Not rsx2PX.EOF 

		PUESTOMASTER = rsx2PX("PUESTOMASTER")
		If rsx2PX("grafica") = "S" Then Exit Do
		Call OFV.cerrarconsulta(rsx2PX)
		StrSql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador,listadepuestos.GRAFICA FROM RelaPuesto  "
		StrSql = StrSql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid) "
		StrSql = StrSql & " WHERE RelaPuesto.puestodepen = 0" & PUESTOMASTER & "  ORDER BY RelaPuesto.puestorel "
		'response.write strsql
		Set rsx2PX = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
             Loop
                    
                    
    Else
        PUESTOMASTER = spuesto
        PUESTOMASTERdescri = ""
    End If

  ''  If xgrafica = "N" Then
        zpuesto = PUESTOMASTER
  ''  Else
   '     zpuesto = spuesto
        PUESTOMASTERdescri = spsector
   ' End If
    zpuesto=puestomaster
    
    '    if     xgrafica = "S" then
    'shref = "OrganigramaPuestosgraf.asp?ID=" & rs(4) & "&Sector=" & OFV.ConvertirCar(rsx2(1), " ", "_")
    shref = "OrganigramaPuestosgraf.asp?ID=" & zpuesto & "&Sector="  & OFV.ConvertirCar(PUESTOMASTERdescri&"", " ", "_")
    
    
    'sinput = "<font face=verdana color=#000080 size=1><b><U>" & rsx2(12) & "</U></b></font>"
    sinput = "<font face=verdana color=#000080 size=1><b><U>" & ecodescri & "</U></b></font>"
    slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

    sinput1 = slink
    '  end if



    '  END IF



    If sinput1 <> "" Then
        sinput = sImageva & sinput1
    End If




    Celda = Celda & OFV.GenCelda("", "uti  ", "", "80%", "", "", "", sinput, "si")

    Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")


    Formulario = Formulario & Fila



    Celda = OFV.GenCelda("", "UTI colspan=2 style='border-top-style:solid;border-top-color:#FFFFFF;border-top-width:1;' ", "", "20%", "", "", "", "&nbsp;", "si")
    Fila = OFV.genrow("", "", "", "", "", "", "", Celda, "si")
    Formulario = Formulario & Fila




    '''''''' REPORTA A:


    '   Formulario = ""



    StrSql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador FROM RelaPuesto  "
    StrSql = StrSql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid) "
    StrSql = StrSql & " WHERE RelaPuesto.puestodepen = 0" & spuesto & "  ORDER BY listadepuestos.nivelj"
    ' response.write strsql &" <br>"

    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, parametros)


    Fila = ""









    posi = 0
    If Not rs.EOF Then

        Celda = OFV.GenCelda("", "UTI   " & wkstylo1, "", "20%", "", "", "", "DEPENDE DE:", "si")

        Do While (Not rs.EOF)

            '' Celda = ""


            'SECTOR
            sinput = "sin definir"
            sinput1 = ""
            sinput2 = ""

            If rs("continuador") = "NO" Then
                sinput1 = rs(3)

                shref = "GENPUESTOFUNCION.asp?PUESTO=" & rs(1) & "&NOMSECTOR=" & OFV.ConvertirCar(rs(4), " ", "_")
                ''''    shref = "OrganigramaPuestosgraf.asp?ID=" & rs(1) & "&Sector=" & OFV.ConvertirCar(rs(3), " ", "_")
                sinput = "<font face=verdana color=#000080 size=1><b><U>" & rs(4) & "</U></b></font>"
                slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

                sinput1 = slink





                If sinput1 <> "" Or sinput2 <> "" Then
                    sinput = sImageva & sinput1 & "&nbsp;-&nbsp;" & sinput2
                End If

                Celda = Celda & OFV.GenCelda("", "uti  ", "", "80%", "", "", "", sinput, "si")

                Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")


                Formulario = Formulario & Fila
            Else
                'SECTOR
                sinput = "sin definir"
                sinput1 = ""
                sinput2 = ""
                spuesto = rs("PUESTOMASTER")
                StrSql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador FROM RelaPuesto  "
                StrSql = StrSql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid) "
                StrSql = StrSql & " WHERE RelaPuesto.puestodepen = 0" & spuesto & "  ORDER BY RelaPuesto.puestorel "
                'RESPONSE.WRITE STRSQL

                Set rsX = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
                If Not rsX.EOF Then
                    Do While Not rsX.EOF


                        spuestox = rsX("PUESTOMASTER")
                        If rsX("continuador") = "NO" Then Exit Do
                        Call OFV.cerrarconsulta(rsX)
                        StrSql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador FROM RelaPuesto  "
                        StrSql = StrSql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid) "
                        StrSql = StrSql & " WHERE RelaPuesto.puestodepen = 0" & spuestox & "  ORDER BY RelaPuesto.puestorel "
                        'response.write strsql
                        Set rsX = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
                    Loop
                    shref = "GENPUESTOFUNCION.asp?PUESTO=" & rsX(1) & "&NOMSECTOR=" & OFV.ConvertirCar(rsX(4), " ", "_")
                    ''''    shref = "OrganigramaPuestosgraf.asp?ID=" & rs(1) & "&Sector=" & OFV.ConvertirCar(rs(3), " ", "_")
                    sinput = "<font face=verdana color=#000080 size=1><b><U>" & rsX(4) & "</U></b></font>"
                    slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

                    sinput1 = slink

                    If sinput1 <> "" Or sinput2 <> "" Then
                        sinput = sImageva & sinput1 & "&nbsp;-&nbsp;" & sinput2
                    End If

                    Celda = Celda & OFV.GenCelda("", "uti  ", "", "80%", "", "", "", sinput, "si")

                    Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")


                    Formulario = Formulario & Fila

                End If

            End If
            posi = posi + 1
            rs.Movenext

            If Not rs.EOF Then
                Celda = OFV.GenCelda("", "UTI   ", "", "20%", "", "", "", "&nbsp;", "si")
            Else
                Celda = OFV.GenCelda("", "UTI colspan=2 style='border-top-style:solid;border-top-color:#FFFFFF;border-top-width:1;' ", "", "20%", "", "", "", "&nbsp;", "si")
                Fila = OFV.genrow("", "", "", "", "", "", "", Celda, "si")
                Formulario = Formulario & Fila
            End If

        Loop



    '    If tipo = "G" Then
      '      shtml = shtml & "<br><center>" & OFV.GenTabla("", "AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ", "", "95%", "", "0", "", "1", "1", "", "", Formulario, "si") & "</center>"
     '   End If

    Else




    End If

   
    sHTML = sHTML & "<br><center>" & ofv.GenTabla("","AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ","","95%","","0","","1","1","","",Formulario,"si") & "</center>"




    '''''''' LE REPORTAN:




    Formulario = ""



    spuesto = request.querystring("puesto")
    StrSql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.CONTINUADOR FROM RelaPuesto  "
    StrSql = StrSql & " left join listadepuestos on  (RelaPuesto.puestoDEPEN = listadepuestos.objid) "
    'Strsql = Strsql & " WHERE listadepuestos.CONTINUADOR='NO' AND RelaPuesto.puestoMASTER = 0" & spuesto & "  ORDER BY RelaPuesto.puestorel "
    StrSql = StrSql & " WHERE  RelaPuesto.puestoMASTER = 0" & spuesto & "  ORDER BY listadepuestos.nivelj "    'RelaPuesto.puestorel "
    ' RESPONSE.WRITE request.querystring & "<br>" & strsql
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, parametros)


    Fila = ""









    posi = 0
  '  If Not rs.EOF And tipo <> "G" Then
  If Not rs.EOF Then

        Celda = OFV.GenCelda("", "UTI   " & wkstylo1, "", "20%", "", "", "", "SUPERVISA A:", "si")

        Do While (Not rs.EOF)


            ''  Celda = ""


            'SECTOR
            sinput = "sin definir"
            sinput1 = ""
            sinput2 = ""

            spuestox = rs(2)
            If rs("continuador") = "NO" Then
                sinput1 = rs(3)

                shref = "GENPUESTOFUNCION.asp?PUESTO=" & rs(2) & "&NOMSECTOR=" & OFV.ConvertirCar(rs(4) & "", " ", "_")
                ''''  shref = "OrganigramaPuestosgraf.asp?ID=" & rs(2) & "&Sector=" & OFV.ConvertirCar(rs(3), " ", "_")
                sinput = "<font face=verdana color=#000080 size=1><b><U>" & rs(4) & "</U></b></font>"
                slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

                sinput1 = slink
            Else
                StrSql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador FROM RelaPuesto  "
                StrSql = StrSql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid) "
                StrSql = StrSql & " WHERE RelaPuesto.puestoMASTER = 0" & spuestox & "  ORDER BY RelaPuesto.puestorel "
                'response.write strsql&"<br>"
                Set rsX = OFV.crearconsultaEx(StrSql, cn, 1, parametros)

                Do While Not rsX.EOF
                    spuestox = rsX("PUESTOdepen")
                    spuestomas = rsX("puestoMASTER")
                    If rsX("continuador") = "NO" Then Exit Do
                    Call OFV.cerrarconsulta(rsX)
                    StrSql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador FROM RelaPuesto  "
                    StrSql = StrSql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid) "
                    StrSql = StrSql & " WHERE RelaPuesto.puestoMASTER = 0" & spuestox & "  ORDER BY RelaPuesto.puestorel "
                    '   response.write strsql&"xxx<br>"
                    Set rsX = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
                Loop
                '   on error resume next
                If rsX.EOF Then
                    ' on error resume next
                    StrSql = " select * from listadepuestos where objid=0" & spuestox
                Else
                    StrSql = " select * from listadepuestos where objid=0" & spuestomas
                End If
                'response.write strsql
                Set rsX = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
                If rsX.EOF Then
                    sinput1 = "sin definir"
                Else
                    sinput1 = rsX(1)
                    '
                    shref = "GENPUESTOFUNCION.asp?PUESTO=" & rsX(0) & "&NOMSECTOR=" & OFV.ConvertirCar(rsX(1) & "", " ", "_")
                    shref = "OrganigramaPuestosgraf.asp?ID=" & rsX(0) & "&Sector=" & OFV.ConvertirCar(rsX(1), " ", "_")
                    sinput = "<font face=verdana color=#000080 size=1><b><U>" & rsX(1) & "</U></b></font>"
                    slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")
                    sinput1 = slink
                End If

            End If




            If sinput1 <> "" Or sinput2 <> "" Then
                sinput = sImageva & sinput1 & "&nbsp;-&nbsp;" & sinput2
            End If

            Celda = Celda & OFV.GenCelda("", "uti  ", "", "80%", "", "", "", sinput, "si")


            Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")


            Formulario = Formulario & Fila



            posi = posi + 1
            rs.Movenext

            If Not rs.EOF Then
                Celda = OFV.GenCelda("", "UTI   ", "", "20%", "", "", "", "&nbsp;", "si")
            Else
                Celda = OFV.GenCelda("", "UTI colspan=2 style='border-top-style:solid;border-top-color:#FFFFFF;border-top-width:1;' ", "", "20%", "", "", "", "&nbsp;", "si")
                Fila = OFV.genrow("", "", "", "", "", "", "", Celda, "si")
                Formulario = Formulario & Fila
            End If

        Loop

        '  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ","","95%","","0","","1","1","","",Formulario,"si") & "</center>"

    Else




    End If


    sinput = sImageIT & "RELACIONES DE AUTORIDAD"

    Celda = OFV.GenCelda("", "UTI  colspan=2 style='color:#FFFFFF;background-color:#888888;' ", "", "", "", "", "", sinput, "si")


    Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")


    Formulario = Fila & Formulario


    If tipo <> "G" Then
        shtml = shtml & "<br><center>" & OFV.GenTabla("", "AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ", "", "95%", "", "0", "", "1", "1", "", "", Formulario, "si") & "</center>"
    End If



    ''''''''''''' DESCRI.PUESTO

    Formulario = ""





    StrSql = "Select * FROM ECOPERFILDSC WHERE "
    StrSql = StrSql & " ECOPERFIL = 0" & sperfil & " ORDER BY SECUEN "
    ''  RESPONSE.WRITE STRSQL & "<BR>"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, parametros)





    posi = 0
    If Not rs.EOF then ' And tipo <> "G" Then

        '  Formulario = Formulario & FILA

        Do While (Not rs.EOF)

            sinput1 = ""

            Fila = ""


            sinput = "sin definir"
            StrSql = "Select * From TIPODSCFUN where OBJID = 0" & rs(3)
            Set rsX2 = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
            If Not rsX2.EOF Then sinput1 = rsX2(1)



            sinput = sImageva & Replace(rs(4) & "", Chr(13), "<br>" & sImageva)

            'Celda = ofv.GenCelda("","UTI   " & wkstylo2,"","20%","","","",sinput1,"si")
            'Celda = CELDA & ofv.GenCelda("","uti  style='color:#000000;' ","","80%","","","",sInput ,"si")
            Celda = OFV.GenCelda("", "uti  colspan=2 style='color:#000000;' ", "", "80%", "", "", "", sinput, "si")
            Fila = Fila & OFV.genrow("", "", "", "", "30", "", "", Celda, "si")

            Formulario = Formulario & Fila

            posi = posi + 1
            rs.Movenext
        Loop



        sinput = sImageIT & "PROPOSITO DEL PUESTO"

        Celda = OFV.GenCelda("", "UTI  colspan=2 style='color:#FFFFFF;background-color:#888888;' ", "", "", "", "", "", sinput, "si")


        Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")


        Formulario = Fila & Formulario


        shtml = shtml & "<br><center>" & OFV.GenTabla("", "AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ", "", "95%", "", "0", "", "1", "1", "", "", Formulario, "si") & "</center>"

    Else



    End If




    ''''''''''' FUNCIONES DEL PUESTO


    Formulario = ""

    StrSql = "Select * FROM ECOPERFILFUN WHERE  tipo <> 'GEN' and "
    StrSql = StrSql & " ECOPERFIL = 0" & sperfil & " ORDER BY TIPO,SECUEN "

    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, parametros)

    Fila = ""





    posi = 0
    If Not rs.EOF  then 'And tipo <> "G" Then

        ' Formulario = Formulario & FILA

        Do While (Not rs.EOF)

            sForm = "formRB" & posi

            sinput1 = ""
            Celda = ""
            Fila = ""

            'ACCION
            sinput = "sin definir"
            StrSql = "Select * From TIPOACCIONFUN where objid = 0" & rs(3)
            Set rsX2 = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
            If Not rsX2.EOF Then sinput1 = rsX2(1)


            'DESCRIPCION
            sinput =posi  +  1 & ".&nbsp;"& sImageva & Replace(rs(4) & "", Chr(13), "<br>" & sImageva)
            Celda = OFV.GenCelda("", "uti  colspan=2 style='color:#000000;' ", "", "80%", "", "", "", sinput, "si")
            Fila = Fila & OFV.genrow("", "", "", "", "30", "", "", Celda, "si")


            Formulario = Formulario & Fila



            posi = posi + 1
            rs.Movenext
        Loop


        sinput = sImageIT & "DESCRIPCION DE FUNCIONES ESPECIFICAS"

        Celda = OFV.GenCelda("", "UTI  colspan=2 style='color:#FFFFFF;background-color:#888888;' ", "", "", "", "", "", sinput, "si")


        Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")
        Formulario = Fila & Formulario


        shtml = shtml & "<br><center>" & OFV.GenTabla("", "AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ", "", "95%", "", "0", "", "1", "1", "", "", Formulario, "si") & "</center>"

    End If




    ' FUNCIONES GENERICAS

    Formulario = ""

    StrSql = "Select * FROM ECOPERFILFUN WHERE TIPO='GEN' AND "
    StrSql = StrSql & " ECOPERFIL = 0" & NIVELJ & " ORDER BY TIPO,SECUEN "
    'RESPONSE.WRITE STRSQL
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, parametros)

    Fila = ""




    posi = 0
    If Not rs.EOF Then

        ' Formulario = Formulario & FILA

        Do While (Not rs.EOF)

            sForm = "formRB" & posi

            sinput1 = ""
            Celda = ""
            Fila = ""

            'ACCION
            sinput = "&nbsp;"
            StrSql = "Select * From TIPOACCIONFUN where objid = 0" & rs(3)
            'RESPONSE.WRITE STRSQL
            Set rsX2 = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
            If Not rsX2.EOF Then sinput1 = rsX2(1)

          '  Celda = OFV.GenCelda("", "uti   style='color:#000000;' ", "", "80%", "", "", "", sinput, "si")
           ' Fila = Fila & OFV.genrow("", "", "", "", "30", "", "", Celda, "si")
             sinput = posi  +  1 & ".&nbsp;"& sImageva & Replace(rs(4) & "", Chr(13), "<br>" & sImageva)
            Celda =  OFV.GenCelda("", "uti   colspan=2 style='color:#000000;' ", "", "80%", "", "", "", sinput, "si")
            Fila = Fila & OFV.genrow("", "", "", "", "30", "", "", Celda, "si")


            Formulario = Formulario & Fila



            posi = posi + 1
            rs.Movenext
        Loop


        sinput = sImageIT & "DESCRIPCION DE FUNCIONES GENERICAS"

        Celda = OFV.GenCelda("", "UTI  colspan=2 style='color:#FFFFFF;background-color:#888888;' ", "", "", "", "", "", sinput, "si")


        Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")



        Formulario = Fila & Formulario

        shtml = shtml & "<br><center>" & OFV.GenTabla("", "AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ", "", "95%", "", "0", "", "1", "1", "", "", Formulario, "si") & "</center>"

    Else



    End If




    Formulario = ""


    StrSql = "Select * FROM ECOPERFILCOM WHERE "
    StrSql = StrSql & " ECOPERFIL = 0" & sperfil & " ORDER BY SECUEN "
    ''  RESPONSE.WRITE STRSQL & "<BR>"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, parametros)

    Fila = ""


    posi = 0
    Formulario = ""
    If Not rs.EOF  then 'And tipo <> "G" Then

        '   Formulario = Formulario & FILA

        Do While (Not rs.EOF)

            Celda = ""
            Fila = ""

            sinput1 = ""
            SCOMPEDAT = "SIN DEFINIR"

            StrSql = "Select * From tipocompetencia where objid = 0" & rs(3)
            '' response.write strsql
            Set rsX2 = OFV.crearconsultaEx(StrSql, cn, 1, parametros)

            If Not rsX2.EOF Then
                SCOMPETIPO = CStr(rsX2(2) & "")
                SCOMPEDAT = CStr(rsX2(1) & "")
            End If

            sinput1 = SCOMPEDAT

            'DESCRIPCION


            If SCOMPETIPO = "1" Then

                SCOMPEDAT = "<FONT COLOR='#993333' >SIN DEFINIR</FONT>"
                SCOMPEDET = rs(4) & ""

                If Not IsNumeric(SCOMPEDET) Then SCOMPEDET = 0

                StrSql = "Select * From TIPOCOMPEDET where objid = 0" & SCOMPEDET
                '' response.write strsql
                Set rsX2 = OFV.crearconsultaEx(StrSql, cn, 1, parametros)

                If Not rsX2.EOF Then
                    SCOMPEDAT = CStr(rsX2(2) & "")
                End If


            Else

                SCOMPEDAT = CStr(rs(4) & "")

            End If


            sinput = sImageva & Replace(SCOMPEDAT & "", Chr(13), "<br>" & sImageva)


            Celda = OFV.GenCelda("", "UTI   " & wkstylo2, "", "20%", "", "", "", sinput1, "si")
            Celda = Celda & OFV.GenCelda("", "uti  style='color:#000000;'  ", "", "80%", "", "", "", sinput, "si")
            Fila = Fila & OFV.genrow("", "", "", "", "30", "", "", Celda, "si")

            Formulario = Formulario & Fila



            posi = posi + 1
            rs.Movenext
        Loop

        sinput = sImageIT & "COMPETENCIAS DEL PUESTO"

        Celda = OFV.GenCelda("", "UTI  colspan=2 style='color:#FFFFFF;background-color:#888888;' ", "", "", "", "", "", sinput, "si")


        Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")


        Formulario = Fila & Formulario


        shtml = shtml & "<br><center>" & OFV.GenTabla("", "AAA style='border-style:solid;border-color:#cccccc;border-width:1;'  ", "", "95%", "", "0", "", "1", "1", "", "", Formulario, "si") & "</center>"

    Else



    End If





    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)


    shtml = shtml & "</BODY>"

    shtml = shtml & "</HTML>"
End If

response.write shtml



%>
