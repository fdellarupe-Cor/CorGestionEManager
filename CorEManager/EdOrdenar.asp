<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

''''''''''''
  Generar = ""
  
  volver = OFV.vueltaasp2("../")
  volver2 = OFV.vueltaasp2("")

  nTipo   = Request.querystring("Tipo")
  nCodigo = Request.querystring("codigo")
  sID     = nCodigo

  SNOMDOC = Request.querystring("NOMDOC")
  SNOMDSC_1 = REPLACE(SNOMDOC,"_"," ") 
  SNOMDSC_2 = REPLACE(SNOMDSC_1,"*"," ") 
  SNOMDSC = REPLACE(SNOMDSC_2,":"," : ") 
  
  
  sTipo   = Request.querystring("XTipo")

  If sNombre = "" Then sNombre = UCase(sTipo)


  vueltadoc  = cstr("../edordenar.asp?Codigo=" & nCodigo & "&Tipo=" & nTipo & "&XTipo=" & sTipo & "&NOMDOC=" & SNOMDOC)

  vueltadoc = REPLACE(vueltadoc,"_","*") 
  vueltadoc = REPLACE(vueltadoc,"&","_") 

  storig = "no"

  If Session("sEstadon") <> "3" Then
     sfont   = "<font id='XXXX' color=#336699 >"
     sffont  = "</font>"
     sfont2  = "<font id='XXXX' color=#993333 >"
     sffont2 = "</font>"
     sfont3  = "<font id='XXXX' color=#777777 >"
     sffont3 = "</font>"
  Else
     sfont   = ""
     sffont  = ""
  End If


  volver2 = volver2 & "?Tipo=" & sTipo & "&Nombre=" & sNombre & "&ID=" & sID & "&Menu="
  volver2 = volver2 & Session("Menu") & "&sitem=" & OFV.convertircar(sitem, " ", "_")
  volver2 = volver2 & "&stitem=" & OFV.convertircar(stitem, " ", "_") & "&sticono=" & sticono
  volver2 = volver2 & "&stnivel=" & stnivel & "&stesq=" & stesq
  If Request.querystring("storig") <> "" Then
    volver2 = volver2 & "&storig=" & Request.querystring("storig")
  End If
  Session("ASP") = volver2

  borderB = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("APLSHW") & ";"
  borderl = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("APLSHW") & ";"

  Session("CBRDICO") = "border-style:none;"

  sImagenM   = "<B><FONT FACE=""corflaticon"" COLOR=#005500 SIZE=3>2</FONT></B>"
  sImagenI2  = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
  sImagenI3  = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM  = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx   = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"

  stnivel = 1
  
    sHTML = ofv.FormHeader(session("FormN")) & chr(13)


  Set cn = OFV.conectar(OFV.strconn2)
    nv = 0
    
    
  sHTML = sHTML &  "<div id='xmt2'  style='height:350px;overflow:auto;" & styl & "' >" & chr(13)
    
  sHTML = sHTML & "<table cellpadding=0 cellspacing=0 width=90% class=aaa>"



      sHTML = sHTML & "<tr>"
      sHTML = sHTML & "<td class=VT colspan=2  style='" & Session("CBRDICO") & "' > "
      'sHTML = sHTML & STIPO
	  'sHTML = sHTML & "Ordenado de esquema:"
      sHTML = sHTML & "</td>"
      'sHTML = sHTML & "<td>" & NCODIGO & "</td>"
	  sHTML = sHTML & "<td>" & "" & "</td>"
      'sHTML = sHTML & "<td colspan=22  >" & NTIPO & "</a></td>"
	  sHTML = sHTML & "<td class=VT colspan=22  style='" & Session("CBRDICO") & "' > " & "Ordenamiento de esquema:" & "</a></td>"
      sHTML = sHTML & "</tr>"

      sHTML = sHTML & "<tr>"
      sHTML = sHTML & "<td>&nbsp;</td>"
      sHTML = sHTML & "</tr>"





 ''''''''' EMPIEZA '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


      Stsup = "Y"



    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td >&nbsp;</td>"



    sxicon = "ICorAbierto.jpg"
    If Trim(sID) = "0" Then sxicon = "ICorRecycle.jpg"

    sxiconv = "<i id='" & snid & "' class='las la-angle-right' color=" & Session("CMFPNTITFT") & "></i>"
    If Trim(sID) = "0" Then
       sxiconv = "<font id='" & snid & "' face='corflaticon' size=3 color='#000000'>&#052;</font>"
    End If

    sHTML = sHTML & "<td class=ttt colspan=2  rowspan=2 style='" & Session("CBRDICO") & "' >"
    sHTML = sHTML & sxiconv
    sHTML = sHTML & "&nbsp;</td>"
    sHTML = sHTML & "<td colspan=22 rowspan=2 class=ttt >" &  "  " & SNOMDSC & "</td>"
    sHTML = sHTML & "</tr>"

    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td>&nbsp;</td>"
    sHTML = sHTML & "<td>&nbsp;</td>"
    sHTML = sHTML & "</tr>"




    StrSql = "SELECT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado as estcod, "
    StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.estado, OBJETOS.INDOBJ FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
    StrSql = StrSql & " where  "
    StrSql = StrSql & "RELAOBJ.codojasoc=OBJETOS.codobjeto and "
    StrSql = StrSql & "OBJPRTY.estado=Estados.objid and "
    StrSql = StrSql & "OBJETOS.codobjeto = OBJPRTY.codobjeto and "
    StrSql = StrSql & "OBJETOS.OBJUN = " & SESSION("UN") & " and "
    StrSql = StrSql & " RELAOBJ.codobjmaster=" & sID
    StrSql = StrSql & " and relaobj.tipobjmaster = '" & NTIPO & "' "
    StrSql = StrSql & " And OBJETOS.objnivel >= " & Session("SegNivel")
    StrSql = StrSql & " And (OBJPRTY.estado=" & Session("sEstadon")
    StrSql = StrSql & " or OBJPRTY.estado=3 "
    If Session("Menu") = "ParaAprobar" Then
       StrSql = StrSql & " or OBJPRTY.estado=1 "
    ElseIf Session("Menu") = "Desarrollo" Then
       StrSql = StrSql & " or OBJPRTY.estado=2 "
    End If
    StrSql = StrSql & " ) "

    StrSql = StrSql & " order by OBJETOS.INDOBJ, OBJETOS.tipobjeto, OBJETOS.descri, OBJETOS.codobjeto "

  '  RESPONSE.WRITE STRSQL & "<BR>"
    sNuevoTipo = "ESQUEMAS"
    nNivel = stnivel + 1

'''''''''''''''''''''''''' TERMINA '''''''''''''''''''''''''''''''''''''''''''''''''''


  Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
  If rs.EOF Then
     If sID <> "xxx" Then
        sHTML = sHTML & "<tr>"
        sHTML = sHTML & "<td colspan=" & ((nNivel - 1) * 2) & " >&nbsp;</td>"
        For xind = 1 To nNivel
        Next
        sHTML = sHTML & "<td colspan=5 >&nbsp;</td><td class=vti colspan=" & (20 - ((nNivel - 1) * 2)) & ">"
        sHTML = sHTML & "No hay Esquemas asociados</td></tr>"
     Else
        sIcono = "ICorRecycle.jpg"
        stitem = "RYS"
        sitem = "CARPETA DE OBJETOS "
        stestado = "N"
        xsf = ""
        xsff = ""
' ---------------------------------



      sHTML = sHTML & "<tr>"


      StrSql = "Select relaobj.relac from relaobj, OBJPRTY, Estados "
      StrSql = StrSql & "where relaobj.codojasoc = OBJPRTY.codobjeto and "
      StrSql = StrSql & "OBJPRTY.estado = Estados.objid and "
      StrSql = StrSql & " relaobj.tipobjmaster = '" & "RYS" & "'"
      StrSql = StrSql & " and relaobj.codobjmaster = 0"
      StrSql = StrSql & " And (OBJPRTY.estado=" & Session("sEstadon")
      StrSql = StrSql & " or OBJPRTY.estado=3 ) "

      Set rsmas = OFV.crearconsultaEx(StrSql, cn, 1, "")
      If Not rsmas.EOF Then
        smas = "<font color=#993333 size=1><b>+</b></font>"
      Else
        smas = "<font color=#993333 size=2><b>-</b></font>"
      End If
      Call OFV.cerrarconsulta(rsmas)

      sHTML = sHTML & "<td colspan=" & ((nNivel - 1) * 2) & "   rowspan=2>&nbsp;</td>"

      For xind = 1 To nNivel
          If xind = nNivel Then

             sHTML = sHTML & "<td >&nbsp;</td>"
             sHTML = sHTML & "<td style='" & borderl & borderb & "'>&nbsp;</td>"

          Else
          End If
      Next

      snid = "0-RYS" & Session("EsqSup")
      snid = snid & "-" & stestado



      sxiconv = "<font id='" & snid & "' face='corflaticon' size=3 color='#f7aa22'>&#052;</font>"


        sHTML = sHTML & "<td class=ttt colspan=2  rowspan=2 style='" & Session("CBRDICO") & "'>"
        sHTML = sHTML & sxiconv
        sHTML = sHTML & "</td>"

        IF xsf <> "" THEN xsf = replace(xsf,"XXXX",snid)

        sHTML = sHTML & "<td rowspan=2>&nbsp;</td>"
        sHTML = sHTML & "<td colspan=" & (20 - ((nNivel - 1) * 2)) & " CLASS=UT  rowspan=2>" & sitem
        sHTML = sHTML & "</td>"
        sHTML = sHTML & "</tr>"


        sHTML = sHTML & "<tr>"
        sHTML = sHTML & "<td >&nbsp;</td>"
        sHTML = sHTML & "<td >&nbsp;</td>"
        sHTML = sHTML & "</tr>"

' ---------------------------------
     End If
  Else

    Do Until rs.EOF



      If UCase(sTipo) = "ESQUEMAS" Then

        If nv = 1 Then
           sHTML = sHTML & "<tr>"

           sHTML = sHTML & "<td >&nbsp;</td>"
           sHTML = sHTML & "<td style='" & borderl & "'>&nbsp;</td>"
           sHTML = sHTML & "</tr>"
        Else
           nv = 1
        End If

        subtprv  = "N"
        stitem   = rs(4)
        stestado = rs(5)
        sNIVE    = "D"

        StrSql = "Select archivo, descri, TIPOCPT From Tipobjeto "
        StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs(4) & "'"
        Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
        If Not rsIcono.EOF Then
          sIcono = rsIcono(0)
          sTipoE = rsIcono(1)
          sNIVE  = rsIcono(2)

          sitem = sTipoE & ": " & rs(1)


        '''  IF sNIVE = "D" THEN  sitem = sitem & sVige

        End If
        Call OFV.cerrarconsulta(rsIcono)
      End If

      If UCase(sTipo) = "SUBTEMAS" Then
         If rs(5) <> "0" Then
            subtprv = "S"
            StrSql = "Select * From subtemaskr "
            StrSql = StrSql & "Where codprodserv = " & rs(0)
            StrSql = StrSql & " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' or grpusr = 'GRP-" & UCase(Session("seggrp")) & "')"
            Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
            If Not rsIcono.EOF Then
               subtprv = "N"
            End If
            Call OFV.cerrarconsulta(rsIcono)
         Else
            subtprv = "N"
         End If

         If subtprv <> "S" Then

            Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs(0) & " AND ATRMTIPOST = 'S' AND ATRMCODES = " & session("UN")
            Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
            If Not rsIcono.EOF Then
               if not rsIcono(4) then  subtprv = "S"
            End If
            Call OFV.cerrarconsulta(rsIcono)

         end if 


         If subtprv <> "S" Then

        If nv = 1 Then
           sHTML = sHTML & "<tr>"

           sHTML = sHTML & "<td >&nbsp;</td>"
           sHTML = sHTML & "<td style='" & borderl & "'>&nbsp;</td>"
           sHTML = sHTML & "</tr>"
        Else
           nv = 1
        End If

            sIcono = rs(4)
            stitem = rs(3)
            sitem = rs(3) & ": " & rs(1)
            stestado = "N"
            xsf = ""
            xsff = ""

         end if

      End If

      If subtprv <> "S" Then

      sForm = "form" & RS(0)

      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=objetos&volver=" & vueltadoc
     sHTML = sHTML & ofv.genform(sForm,"",sAccion,"","post","",Fila,"no")

      sHTML = sHTML & "<tr>"


      StrSql = "Select relaobj.relac from relaobj, OBJPRTY, Estados "
      StrSql = StrSql & "where relaobj.codojasoc = OBJPRTY.codobjeto and "
      StrSql = StrSql & " OBJPRTY.estado = Estados.objid and "
      StrSql = StrSql & " relaobj.tipobjmaster = '" & rs(2) & "'"
      StrSql = StrSql & " and relaobj.codobjmaster = " & rs(0)
      StrSql = StrSql & " And (OBJPRTY.estado=" & Session("sEstadon")
      StrSql = StrSql & " or OBJPRTY.estado=3 "
      If Session("Menu") = "ParaAprobar" Then
         StrSql = StrSql & " or OBJPRTY.estado=1 "
      ElseIf Session("Menu") = "Desarrollo" Then
         StrSql = StrSql & " or OBJPRTY.estado=2 "
      End If
      StrSql = StrSql & " ) "



      Set rsmas = OFV.crearconsultaEx(StrSql, cn, 1, "")
      If Not rsmas.EOF Then
        smas = "<font color=#993333 size=1> (+)</font>"
      Else
        smas = ""
      End If
      Call OFV.cerrarconsulta(rsmas)

      If ((nNivel - 1) * 2) > 0 Then
         sHTML = sHTML & "<td colspan=" & ((nNivel - 1) * 2) & "  rowspan=2 >&nbsp;</td>"
      End If

      For xind = 1 To nNivel
        If xind = nNivel Then

             sHTML = sHTML & "<td  >&nbsp;</td>"
             sHTML = sHTML & "<td style='" & borderl & borderb & "'>&nbsp;</td>"

        Else
        End If
      Next

    snid = rs(0) & "-" & rs(2) & Session("EsqSup")
    snid = snid  & "-" & stestado


    if  sNIVE = "D" then

        If CStr(stestado) = "3" Then
           'sxiconv = "<font id='" & snid & "' face='corflaticon' size=4 color='#99CCEE'>&#0157;</font>"
		   sxiconv = "<i id='" & snid & "' class='las la-file-alt' ></i>"
           sVige = ""
        ElseIf CStr(stestado) = "1"  Then
           sxiconv = "<font id='" & snid & "' face='corflaticon' size=4 color='#00AA00'>&#0157;</font>"
           sVige = " (" & rs(3) & ")"
        ElseIf CStr(stestado) = "2"  Then
           sxiconv = "<font id='" & snid & "' face='corflaticon' size=4 color='#993333'>&#0157;</font>"
           sVige = " (" & rs(3) & ")"
        Else
           sxiconv = "<font id='" & snid & "' face='corflaticon' size=4 color='#000000'>&#0157;</font>"
           sVige = " (" & rs(3) & ")"
        End If


    else

        'sxiconv = "<font id='" & snid & "' face='corflaticon' size=4 color='#336699'>&#0204;</font>"
		sxiconv = "<i id='" & snid & "' class='las la-folder' ></i>"

    end if


        sHTML = sHTML & "<td class=ttt colspan=2  rowspan=2 style='" & Session("CBRDICO") & "'>"
        sHTML = sHTML & sxiconv & "&nbsp;"
        sHTML = sHTML & "</td>"


       if  sNIVE <> "D" then

           sitemfl = "&nbsp;<u >" & sitem & "</u>"  & smas
       else
           sitem   = sitem & sVige
           sitemfl = "&nbsp;" & sitem & smas

       end if

    'Orden
      sInput = ofv.GenerarInput(rs(6).name,rs(6),"text",sform,"5","dtn","")


      sHTML = sHTML & "<td   rowspan=2>&nbsp;</td>"
      sHTML = sHTML & "<td CLASS=UT colspan=" & (19 - ((nNivel - 1) * 2)) & "   rowspan=2>" 
      sHTML = sHTML & sitemfl
      sHTML = sHTML & "</td>"
      sHTML = sHTML & "<td   rowspan=2>" & sInput & "</td>"
      sHTML = sHTML & "</tr>"

     sHTML = sHTML & ofv.genfform()

      End If

      rs.movenext
    Loop
  End If

  Call OFV.cerrarconsulta(rs)
  Call OFV.cerrarconn(cn)

        If nv = 1 Then
           sHTML = sHTML & "<tr>"
           sHTML = sHTML & "<td >&nbsp;</td>"
           sHTML = sHTML & "<td >&nbsp;</td>"
           sHTML = sHTML & "</tr>"
        End If

      sHTML = sHTML & "<tr>"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=1% >"
      sHTML = sHTML & "<td width=76% >"
      sHTML = sHTML & "<td width=1% >"

      sHTML = sHTML & "<tr>"

  sHTML = sHTML & "</table>"
  
  sHTML = sHTML & "</div>"
  
  sHTML = sHTML & "</body>"

  sHTML = sHTML & "</html>"
  
  Generar = sHTML
    




'''''''''''

  sHTML = Generar





end if

response.write sHTML

%>
