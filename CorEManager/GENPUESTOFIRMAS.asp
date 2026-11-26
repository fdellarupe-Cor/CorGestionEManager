
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


    shtml = OFV.MenuHeader("", "")  & CorPnlFnc("N","")

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

    StrSql = StrSql & " where "
    StrSql = StrSql & "Ecoperfil.codsector = listadepuestos.objid and Ecoperfil.codsector= '" & spuesto & "'  and Ecoperfil.codperfil = '****PUESTO' "
   ' response.write strsql
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, parametros)



    If rs.EOF Then
        ecodescri = "SIN DEFINIR"

        Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", "&nbsp;", "si")
        Fila = Fila & OFV.genrow("", "", "", "", "20", "", "", Celda, "si")

        Celda = OFV.GenCelda("", "uti colspan=2 style='border-style:solid;border-color:#888888;border-width:1;'  ", "center", "", "", "", "", "<big>No hay regimen de firmas  para el puesto.</big>", "si")
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

  '  sinput = sImageIT & "RELACIONES DE AUTORIDAD"
  '
  '   celda = OFV.GenCelda("", "UTI  colspan=2 style='color:#FFFFFF;background-color:#888888;' ", "", "", "", "", "", sinput, "si")
  '
  '
  '
  ''    Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")

  '   shtml = shtml & "<br><center>" & OFV.GenTabla("", "AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ", "", "95%", "", "0", "", "1", "1", "", "", fila, "si") & "</center>"

     BTIT1 = "<A CLASS=CC HREF='BSQINFO.ASP?BSQ=PORPUESTO&ID="&spuesto&"'>Ver Datos de los Ocupantes del Puesto</A>"
     BTIT2 = "<A CLASS=CC HREF='GENPUESTOFUNCION.ASP?PUESTO="&spuesto &"'   >VER LAS FUNCIONES DEL PUESTO&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</A>"

     Celda = OFV.GenCelda("", "UTI   " & wkstylo1, "", "", "", "", "", BTIT1, "si")
     Celda = CELDA & OFV.GenCelda("", "UTI   " & wkstylo1, "", "", "", "", "", BTIT2, "si")

     Fila = OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")

    shtml = shtml & "<BR><center>" & OFV.GenTabla("", "AAA style='border-style:solid;border-color:#cccccc;border-width:1;background-color:#888888;' ", "", "", "", "0", "", "1", "1", "", "", Fila, "si") & "</center>"





    Formulario = ""
 ' Fila=""




    StrSql = "Select * FROM ecoperfilFIR WHERE "
    StrSql = StrSql & " ECOPERFIL = 0" & sperfil & " ORDER BY SECUEN "
    ''  RESPONSE.WRITE STRSQL & "<BR>"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, parametros)





    posi = 0
    If Not rs.EOF then

        '  Formulario = Formulario & FILA

        Do While (Not rs.EOF)

            sinput1 = ""

            Fila = ""


            sinput = "sin definir"
            StrSql = "Select * From REGIMENFIRMAS where OBJID = 0" & rs(3)
            'RESPONSE.WRITE STRSQL
            Set rsX2 = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
            If Not rsX2.EOF Then
            sinput1 = rsX2(1)



            sinput = sImageva & Replace(rs(4) & "", Chr(13), "<br>" & sImageva)


            Celda = OFV.GenCelda("", "xxx ", "", "", "", "", "", genlnkfirmas(rsX2("rellink"),sinput1) , "si")
            Celda = CELDA & OFV.GenCelda("", "uti", "", "", "", "", "", sinput, "si")
            Fila = Fila & OFV.genrow("", "", "", "", "", "", "", Celda, "si")
            end if

            Formulario = Formulario & Fila

            posi = posi + 1
            rs.Movenext
        Loop



        sinput = sImageIT & "REGIMEN DE FIRMAS"

        Celda = OFV.GenCelda("", "UTI  colspan=2 style='color:#FFFFFF;background-color:#888888;' ", "", "", "", "", "", sinput, "si")


        Fila = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")


        Formulario = Fila & Formulario


        shtml = shtml & "<br><center>" & OFV.GenTabla("", "AAA   style='border-style:solid;border-color:#cccccc;border-width:1;' ", "", "95%", "", "0", "", "1", "1", "", "", Formulario, "si") & "</center>"

    Else



    End If









    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)


    shtml = shtml & "</BODY>"

    shtml = shtml & "</HTML>"
End If



function genlnkfirmas(id,xtexto)
dim descri
        descri=""
        xsstyle = " ut "
             wstextoLNK = id

             if not isnull(wstextoLNK) AND LEN(wstextoLNK) > 0 and wstextoLNK <> " " then

                if wstextoLNK <> "0" then

             wlink  = clng(wstextoLNK)
             strsql = "SELECT *  From objlnk where objid = " & wlink
        ' response.write strsql
             set rst = ofv.crearconsultaEx(StrSql,cn,1,volver)
              xdisplay=""
             if not rst.eof then
                xdisplay=" style={display:none;visibility:hidden;}"
                wasotxt     = rst(14) & ""
                wasotxtstr  = CLNG("0" & rst(15))
                wasotxtcan  = CLNG("0" & rst(16))
                wasotxtdat  = rst(17) & ""

                IF wasotxt = "" THEN wasotxt = "N"



                wNV1    = rst(11) & ""
                wNV2    = rst(12) & ""
                wNV3    = rst(13) & ""

                IF  WNV1 = "" THEN WNV1 = "0"
                IF  WNV2 = "" THEN WNV2 = "0"
                IF  WNV3 = "" THEN WNV3 = "0"

                codobjetoL = clng(rst("codobjetoL"))
                wtipoLc    = rst(7)
                if trim(rst("Ltipo")) = "L" then

                   wUrlDesc = rst(9) & ""


                   wNomL    = wUrlDesc
                   if wNomL = "" then
                      wNomL = "S/D"


                      Strsql = "Select * From csGenObjetos where estado <> 4 and codobjeto = " & codobjetoL

                      set rst2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                      if not rst2.eof then
                         wNomL = rst2(1)
                      else
                         wNomL =""' "S/D"
                      end if
                      call ofv.cerrarconsulta(rst2)

                   end if

               '''  sdestino = FNCODOBJ & ""


                 if cstr(codobjetoL & "")  = cstr(FNCODOBJ & "") then

	           sHref = "BarraSol.asp?Tipo=" & FNTIPOBJ & "&Codigo=" & FNCODOBJ & "&Menu=&Sector=&MAINAPL=EM&Funcion=CD&SSCLICK=" & WNV1 & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3 & "&NIV4=0&NIV5=0"

	           sFuncion = " "
	           sTarget  = "WK1"

                  ' STab (yyy,zzz,8)=sHref

                 else

                   sActcon = "Consultar.asp?Codigo=" & codobjetoL & "&Tipo=" & wtipoLc
                   sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
                   sActcon = sActcon & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

	           sTarget   = ""

	           sHref    = "&nbsp;"
	     ' response.write sActcon
	           sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoLc & "'); return false;" & Chr(34)

                 end if

	           slnktit  = wNomL
                   wNomL    = pStyle & "&nbsp;<u><i>" & wNomL & "</i></u>&nbsp;" & pStyleF
	           sImagen  = wNomL



	           sLink    = OFV.GenLink("", xsstyle & " title='" & slnktit & "' ",sHref, sTarget, "", sFuncion, "NNNNNN", "si")
                   descri   = sLink
                   call ofv.cerrarconsulta(rst2)
                  ' STab (yyy,zzz,8)="AbrirCons('" & sActcon & "','" & wtipoLc & "'); return false;"

               else

                   if  not rst.eof then

                       wUrl  = rst("Lurl") & ""

                       if not isnull(wUrl) then
                          sHrefu = instr(wUrl,":")
                          if sHrefu > 0 then
                             sHrefu = wUrl
                          else
                             sHrefu = "HTTP://"  & wUrl
                          end if
                       else
                          sHrefu = "javascript:void(0);"
                       end if

	               slnktit  = rst("LurlDesc")
                       wNomL    = pStyle & "&nbsp;<u><i>" & rst("LurlDesc") & "</i></u>&nbsp;" & pStyleF
	               sImagen  = wNomL

                       descri = ofv.GenLink("",xsstyle & " title='" & slnktit & "' ",GenKeyLnks(sHrefu),"_blank","","","NNNNNN","si")
                  '     STab (yyy,zzz,8)=GenKeyLnks(sHrefu)
                   end if

               end if


           end if
           call ofv.cerrarconsulta(rst)

         end if

      end if

wasotxtstr = 0
wasotxt = "N"

     IF wasotxt = "S" THEN

      if len(sTextoV) > 0 then

         if  wasotxtstr > 0  and wasotxtcan > 0 then

            sTextoV = left(sTextoV,wasotxtstr - 1) & "" & replace(descri,"NNNNNN",pStyle & "&nbsp;<u><b style='color:" & Session("APLFT") & "' >" & mid(sTextoV,wasotxtstr,wasotxtcan) & "</b></u>&nbsp;" & pStyleF) & "" & right(sTextoV,len(sTextoV) - (wasotxtstr + wasotxtcan - 1))

         end if

         sTextoV =  replace(sTextoV,chr(13)," <br>")

      end if


     end if


      IF wasotxt = "N" THEN

         if len(sTextoV) > 0 then

             'sTextoV = supersec & "---" & ta& replace(sTextoV,chr(13)," <br>")
             sTextoV =  replace(sTextoV,chr(13)," <br>")

         end if

       '''  if  sImagen <> "NO" then sTextoV = sTextoV & replace(descri,"NNNNNN",sImagen)


            sTextoV = replace(descri,"NNNNNN",pStyle & sTextoV & pStyleF)



      end if




       ' if STab (yyy,zzz,3) <> "" then
'
       '    slink = "<img  src=" & chr(34) & GenKeyLnks("estructura/" & STab (yyy,zzz,3))  & chr(34) & " "
       '    slink = slink & sreduc & sbord & " >"
       '    slink2 = slink
'
       '  ''  slink = "<a href='" & GenKeyLnks("estructura/" & STab (yyy,zzz,3))  &  "' target='_blank' >" & slink2 & "</a>"
'
'
       '    slink = replace(descri,"NNNNNN",slink2)
       '
'
       ' else
       '    slink = ""
       ' end if

       '   slink = ""

          slinkx = replace(descri,"NNNNNN","<u><b>" & xtexto & "</b></u>")
       '    slinkx = replace(slinkx,"<A ","<A id=" & zzz & "_" & yyy )
genlnkfirmas= slinkx

end function

response.write shtml



%>
