<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0
  server.scripttimeout = 600
  response.buffer = true



''''''''''''''


  XXXX = NEWEMPRE("NO")

    volver = OFV.vueltaasp2("../")
    volver2 = OFV.vueltaasp2("")

    sNombre = "Tema - Subtema"

    sID = Request.querystring("ID")

    stexto = Request.querystring("texto")


    sitem = OFV.convertircar(Request.querystring("sitem"), "_", " ")
    stitem = OFV.convertircar(Request.querystring("stitem"), "_", " ")
    Stnivel = Request.querystring("stnivel")
    If Request.querystring("stesq") = "" Then
        stesq = "TEMA"
    Else
        stesq = Request.querystring("stesq")
    End If

    volver2 = volver2 & "?ID=" & sID
    volver2 = volver2 & "&sitem=" & OFV.convertircar(sitem, " ", "_")
    volver2 = volver2 & "&stitem=" & OFV.convertircar(stitem, " ", "_")
    volver2 = volver2 & "&stnivel=" & Stnivel & "&stesq=" & stesq
    volver2 = OFV.convertircar(volver2, "&", "*")

    ReDim Preserve bConsultar(4)
    OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "4"
    bConsultar(4) = OFV.AConsultar
    OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "2"
    bConsultar(2) = OFV.AConsultar
    OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "1"
    bConsultar(1) = OFV.AConsultar
    OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "3"
    bConsultar(3) = OFV.AConsultar

    sVerEstado = "  and Objetos.objnivel >= " & Session("SegNivel") & " "




    sverestado = sverestado & " and Objprty.Estado In (9,"
    For i = 1 To 4
        If bConsultar(i) Then sVerEstado = sVerEstado & i & ","
    Next
    sVerEstado = Left(sVerEstado, Len(sVerEstado) - 1) & ") "


  sHTML      = ofv.FormHeader(session("FormN"))



    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,status=yes,height=570,width=790,resizable=yes" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13)

    xs = xs & "function Axn(c1,c2,c3) {" & Chr(13)

    xs = xs & " var mydat= 'Consultar.asp?DOCMODAL=CL&Codigo=' + c1 + '&Tipo=' + c2 + '&Perfil=' + c3 + '&Datos=&Volver=" & volver2 & "'; "

    xs = xs & " myWin= open(mydat"
    xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,status=yes,height=570,width=790,resizable=yes" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13)


    xs = xs & "</script>" & Chr(13)

    sHTML = sHTML & xs

  sImagenM   = "<B><FONT FACE=""corflaticon"" COLOR=#005500 SIZE=3>2</FONT></B>"
  sImagenI2  = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
  sImagenI   = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI3  = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM  = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx   = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"


  Fila    = ""
  Celda   = ""
  sHref   = ""

  dim tabprocesos (100,200,20)
  dim tabprocord  (100,5)


  ttabprocesos = 0


  Fila = ""


   sHTML = sHTML & "<center>"


    Response.write sHTML
    sHTML = ""
    Response.flush

    Set cn = OFV.conectar(OFV.strconn2)
    
    xTrColor = "#FFFFFF"     
   


    '  sHTML = sHTML & "<br>"

    '  sHTML = sHTML  &  "<DIV id='DTGENERAL' STYLE='display: compact;' >"



'''''''''''''' buscador por texto     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  TABLAB = ""
  TABLAB = TablaB & "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:NONE;height:50px;border-style:solid;border-width:2;border-color:#ffffff;' >"

  sInputT  = ofv.GenerarInput("textot",stexto,"agregar","","70","dt style='border-style:solid;border-width:1;border-color:#666666;background-color:#f9f9f9;color:#000000;font-size:9pt;' ","")

    celda  = ""
    Celda      = Celda & ofv.GenCelda("","vti STYLE='cursor:pointer;color:#990000;font-size:9pt;' ","","30%","","","","Buscador","si")
    Celda      = Celda & ofv.GenCelda("","UT  ","right","30%","","","",sInputT,"si")

    sInput   = ofv.GenerarInput("","Buscar","submit","","5","at STYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;' ","")
    Celda    = Celda & ofv.GenCelda("","vt","","40%","","","",sInput,"si")

    formx    = ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","","","",Celda,"si")

    sTARGET  = ""
    sAccion  = "consBuscadorSet.asp"

    FILA     = ofv.GenForm("Subconsulta","",sAccion,STARGET,"post","",formx,"si")

  TABLAB = TablaB & ofv.GenTabla("","aaa STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ","","98%","","0","","-2","0","","",Fila,"si")

  TABLAB = TablaB & "</DIV>"

           sHTML = sHTML  &   TABLAB


       xsclase = " CLASS=VTI STYLE='background-color:#f0eedd;color:#000000;font-size:9pt;' "



      XTABLAN = ""


      iproc   = 0 


      iprdoc   = 0 

      dFechai = clng(date()) - clng(cdate("01-01-1900"))


    if session("objtext") = "" then

       sVino = "N"



    else

    tabobjetos = split(session("objtext"),chr(9)) 
    wspsep = ""
    wspbet = "" 

    for ii=0 to ubound(tabobjetos)
        wspbet = wspbet & wspsep &  "'" & tabobjetos(ii) & "'"
        wspsep = ","
    next


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


       Strsql = "select a.cobj, B.DESCRI, b.tipobjeto, e.estado,"
       Strsql = Strsql & " b.tipobjeto, c.fecemis, c.fecvigen, c.fecbaja, "
       StrSql = StrSql & " c.autor, c.aprobo, c.elimino, c.fecvto, b.OBJUN, b.indobj, "
       Strsql = Strsql & "  f.descri as tobjdescri , f.perfil, f.TIPOCPT, g.objcopete,    "
       Strsql = Strsql & "  c.fecpublic, g.objimagen, g.objurldsc, t.descri, st.descri   "
       Strsql = Strsql & " from csrelpath a "
       Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto) "
       Strsql = Strsql & " inner join objprty c on (a.cobj = c.codobjeto) "
       Strsql = Strsql & " inner join estados e on (c.estado = e.objid) "
       Strsql = Strsql & " inner join tipobjeto f on (b.tipobjeto = f.tipobjeto) "
       Strsql = Strsql & " inner join tema      t on (a.tema = t.codtema) "
       Strsql = Strsql & " inner join prodserv  st on (a.stema = st.codprodserv) "
       Strsql = Strsql & " left  join objinfo g on (b.codobjeto = g.codobjeto) "

       StrSql = StrSql & " where "

       StrSql = StrSql & "  a.cobj in (" & wspbet & ") "


       StrSql = StrSql & " order by b.tipobjeto, b.indobj,  b.descri, e.estado "


    
    sNuevoTipo = "ESQUEMAS"
    nNivel = Stnivel + 1
    Session("sID2") = sID
 '''   response.write strsql & "<br>" & chr(13)
    sVino = "N"

    tipoant = "NO"
    canttipo = 0

    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If rs.EOF Then

       sVino = "N"


    Else



        sVino = "N"

        nlinea = 0
        Do Until rs.EOF
             sVinook = "N"
             if  cstr(rs(12) & "") = cstr(session("UN") & "") then
   '' response.write "<br><br>entro"
                 sVino   = "S"
                 sVinook = "S"
             end if 


          if  sVinook = "S" then

                stitem  = rs(4)
                sTipoE  = rs(14)
                sPerfil = rs(15)
                sNIVE   = rs(16)

                sitem = Capitalize(rs(1))

            if tipoant <> sTipoE then

               wspleftb = ""

               if tipoant <> "NO" then 


               else

                  wspleftb = "border-left-style:solid;border-left-width:1;border-left-color:#aa0000;"  

               end if

               tipoant = sTipoE

               iproc = iproc + 1

               tabprocesos (iproc,0,0) = sTipoE
               tabprocesos (iproc,0,1) = 0


               tabprocord  (iproc,0)   = iproc
               tabprocord  (iproc,1)   = 0

               ttabprocesos = iproc

               iprdoc   = 0 


            end if 




               iprdoc = iprdoc + 1

               tabprocesos (iproc,0,1) = iprdoc

               tabprocord  (iproc,1)   = iprdoc


               tabprocesos (iproc,iprdoc,1) = sitem                        'nombre
               tabprocesos (iproc,iprdoc,2) = rs(0)                        'codigo
               tabprocesos (iproc,iprdoc,3) = sPerfil                      'perfil 
               wspwkcopete = rs(17) & ""

               if wspwkcopete = "" then
                  tabprocesos (iproc,iprdoc,4) = "Documento normativo"        'copete
               else
                  tabprocesos (iproc,iprdoc,4) = Capitalize(ltrim(wspwkcopete))      'copete
               end if



               tabprocesos (iproc,iprdoc,5) = rs(6)                        'fec.vigencia


               If clng("0" & rs(12)) = 1 and  SESSION("UN") <> 1 and  SESSION("EMPMUN") = "S" Then
                  tabprocesos (iproc,iprdoc,6) = 1                            'corporativo 0-no 1-si
               Else
                  tabprocesos (iproc,iprdoc,6) = 0                            'corporativo 0-no 1-si
               End If

               tabprocesos (iproc,iprdoc,7) = stitem                       'stitem
               tabprocesos (iproc,iprdoc,8) = rs(18)                        'fec.public
               tabprocesos (iproc,iprdoc,9) = rs(19)                        'objimagen
               tabprocesos (iproc,iprdoc,10) = rs(20)                        'objurldsc
               tabprocesos (iproc,iprdoc,11) = rs(21)                        'tema
               tabprocesos (iproc,iprdoc,12) = rs(22)                        'subtema



            end if

            rs.movenext
            nlinea = nlinea + 1
        Loop





    End If

    end if




    if  sVino = "N" then

        Celda   = ""

        sLink   = "&nbsp;"
        Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

        sLink   = "&nbsp;"
        Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

        sLink   = "No se encuentran documentos para esta busqueda."
        Celda   = Celda & ofv.GenCelda("","vti STYLE='cursor:pointer;color:#990000;font-size:9pt;' ","","90%","","","",sLink,"si")

        Fila    = ofv.GenRow("","aaa bgcolor='" & xTrColor & "' ","","","25","","",Celda,"si")

        TABLAc  = ofv.GenTabla("","aaa STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ","","98%","20%","0","","-2","0","","",Fila,"si")

        sHTML = sHTML & TABLAc

    else 


     ''''   response.write "cant proc: " & ttabprocesos & "<br>"


       if ttabprocesos > 1 then

          for ii=1 to ttabprocesos

              for ii2=1 to (ttabprocesos - ii)

                  if ii2 < ttabprocesos then
                     if (tabprocord  (ii2,1) < tabprocord  (ii2+1,1)) or ( (tabprocord  (ii2,1) = tabprocord  (ii2+1,1)) and (tabprocord  (ii2,0) > tabprocord  (ii2+1,0)) ) then

                        wspwkind =  tabprocord  (ii2,0)
                        wspwkpro =  tabprocord  (ii2,1)

                        tabprocord  (ii2,0) = tabprocord  (ii2+1,0)
                        tabprocord  (ii2,1) = tabprocord  (ii2+1,1)

                        tabprocord  (ii2+1,0) = wspwkind
                        tabprocord  (ii2+1,1) = wspwkpro



                     end if

                  end if


              next


          next

       end if


        for ii=1 to ttabprocesos

               iproc = tabprocord  (ii,0)


          ''     tabprocesos (iproc,0,0)  '' descripcion del tipo de documento


          ''     tabprocesos (iproc,0,1)  '' cantidad de documentos 


            XTABLAN = XTablaN & chr(13)

            XTABLAN = XTablaN & "<div id='DTNOV" & ii & "' style='float: left; display: inline ;  "


           XTABLAN = XTablaN & " width:30%;"

           XTABLAN = XTablaN & "border-right-style:solid;border-right-width:1;border-right-color:#dddddd;' >"


           XTABLAN = XTablaN &  "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
           XTABLAN = XTablaN &  "<tr HEIGHT=25  >"
           XTABLAN = XTablaN &  "<td " & xsclase & "   >" & Capitalize(tabprocesos (iproc,0,0)) & "</TD>"
           XTABLAN = XTablaN &  "</TR>"
           XTABLAN = XTablaN &  "</table>"

           XTABLAN = XTablaN & "<div id='DTNOVINT" & ii & "' style='"

           XTABLAN = XTablaN & "border-style:solid;border-width:1;border-color:#eeeeee;' >"
 
           XTABLAN = XTablaN &  "<table width=99% cellspacing=0 cellpadding=0  >"


           for iprdoc=1 to tabprocesos (iproc,0,1)


            ''   tabprocesos (iproc,iprdoc,1) = sitem                        'nombre
            ''   tabprocesos (iproc,iprdoc,2) = rs(0)                        'codigo
            ''   tabprocesos (iproc,iprdoc,3) = sPerfil                      'perfil 
            ''   tabprocesos (iproc,iprdoc,4) = "Documento normativo"        'copete
            ''   tabprocesos (iproc,iprdoc,5) = rs(6)                        'fec.vigencia
            ''   tabprocesos (iproc,iprdoc,6) = 0                            'corporativo 0-no 1-si
            ''   tabprocesos (iproc,iprdoc,7) = stitem                       'stitem
            ''   tabprocesos (iproc,iprdoc,9) = rs(19)                        'objimagen
            ''   tabprocesos (iproc,iprdoc,10) = rs(20)                        'objurldsc
            ''   tabprocesos (iproc,iprdoc,11) = rs(21)                        'tema
            ''   tabprocesos (iproc,iprdoc,12) = rs(22)                        'subtema


                 xtmsgIMG = tabprocesos (iproc,iprdoc,9) & ""
                 xtmsgVID = tabprocesos (iproc,iprdoc,10) & ""
                 xtmcod   = tabprocesos (iproc,iprdoc,2) & ""

                 xtmdtema   = tabprocesos (iproc,iprdoc,11) & ""
                 xtmdstema  = tabprocesos (iproc,iprdoc,12) & ""



                If tabprocesos (iproc,iprdoc,6) = 1 Then
                   smasc = "&nbsp;&nbsp;&nbsp;<font color=#993333 ><font class=btic title='Corporativo' >G</font></font>"
                Else
                   smasc = "&nbsp;"
                End If


                Celda   = ""

                sImage  = "<font face='corflaticon' size=4 color='#005500' >&#0157;</font>"
                sHref    = "&nbsp;"
                sFuncion = " onclick=" & Chr(34) & " Axn('" & tabprocesos (iproc,iprdoc,2) & "','" & tabprocesos (iproc,iprdoc,7) & "','" & tabprocesos (iproc,iprdoc,3) & "'); return false;" & Chr(34)

                sLink1    = "&nbsp;<font color='#AA0000' size=2 >&#0149;</font>&nbsp;"
                sLink    = "&nbsp;&nbsp;" & tabprocesos (iproc,iprdoc,1)




                       '  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                       '  XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                       '  XTABLAN = XTablaN &  "</TR>"



                         xtmsgtit = tabprocesos (iproc,iprdoc,1)


                         if (dFechai - tabprocesos (iproc,iprdoc,8)) < 31 then


                            xtmsgtit = "<B><font color='#ffffff' STYLE='background-color:#ff0000;' size=1 >&nbsp;nuevo&nbsp;</font></B>&nbsp;&nbsp;" & xtmsgtit

                         end if


                         XTABLAN = XTablaN &  "<tr HEIGHT='32' bgcolor='#f0f0f0' >"
                         XTABLAN = XTablaN &  "<td class=Vti colspan=8 STYLE='cursor:pointer;color:#990000;font-size:9pt;"
                         XTABLAN = XTablaN &  "border-top-style:solid;border-top-width:3;border-top-color:#ffffff;' "
                         XTABLAN = XTablaN &  sFuncion & " title='Consultar.' >" & xtmsgtit & "</TD>"
                         XTABLAN = XTablaN &  "</TR>"

                         XTABLAN = XTablaN &  "<tr >"
                             XTABLAN = XTablaN &  "<td class=UT width='5%'  >&nbsp;</TD>"
                         for yy=1 to 6
                             XTABLAN = XTablaN &  "<td class=UT width='15%'  >&nbsp;</TD>"
                         next
                             XTABLAN = XTablaN &  "<td class=UT width='5%'  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                    IF (xtmsgIMG & "") = "1" THEN


                    Strsql = "Select * From objinfoIMG  where codobjeto = 0" & xtmcod & " ORDER BY objSECUEN, OBJID"
                    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                    IF NOT RS1.EOF THEN



                         XTABLAN = XTablaN &  "<tr bgcolor='#F0F0F0' >"
                         XTABLAN = XTablaN &  "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"

                             XTABLAN = XTablaN &  "<table cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#990000;' >"
                             XTABLAN = XTablaN &  "<tr   >"


                         XSFSTYLE= " STYLE='display:block;' "

                       DO WHILE NOT rs1.eof 

                             XTABLAN = XTablaN &  "<td CLASS=UT " & XSFSTYLE & " ><IMG SRC='" & RS1(6) & "' ></TD>"
                             XSFSTYLE= " STYLE='display:none;' "


                        RS1.MOVENEXT
                       LOOP

                             XTABLAN = XTablaN &  "</TR>"
                             XTABLAN = XTablaN &  "</table>"

                         XTABLAN = XTablaN &  "</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                    END IF






                    ELSEIF (xtmsgIMG & "") = "2" THEN


                         XTABLAN = XTablaN &  "<tr bgcolor='#F0F0F0' >"
                         XTABLAN = XTablaN &  "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"
                         XTABLAN = XTablaN &  xtmsgVID
                      '   XTABLAN = XTablaN &  REPLACE(xtmsgVID & "",CHR(34),CHR(34) & CHR(34))
                         XTABLAN = XTablaN &  "</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"

                    END IF


               ''  xtmdtema   = tabprocesos (iproc,iprdoc,11) & ""
               ''  xtmdstema  = tabprocesos (iproc,iprdoc,12) & ""


                         xtmsgCOP = tabprocesos (iproc,iprdoc,4)


                         XTMSGSEND = "<B><font color=#888888 >" & gtfechadorsh(tabprocesos (iproc,iprdoc,5)) & "</font></B>"
                         XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
                         XTMSGSEND = XTMSGSEND & "<font color=#000000 style='font-weight:bold;' >&nbsp;" & xtmdtema & "&nbsp;-&nbsp;" & xtmdstema & "</font>"
                         XTMSGSEND = XTMSGSEND & "<br>"
                         XTMSGSEND = XTMSGSEND & "<font color=#000000 style='font-size:8pt;font-weight:normal;' >" & xtmsgCOP & "</font>"


                         XTABLAN = XTablaN &  "<tr  >"
                         XTABLAN = XTablaN &  "<td class=ut VALIGN='TOP' align='center' ><font color='#AA0000' size=2 >&#0149;</font></TD>"
                         XTABLAN = XTablaN &  "<td class=ut colspan=6 STYLE='color:#000000;text-align:justify;' >"
                         XTABLAN = XTablaN &   XTMSGSEND & "</TD>"
                         XTABLAN = XTablaN &  "<td class=ut  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8 STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;' >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"



           next


           XTABLAN = XTablaN & "</table>"
           XTABLAN = XTablaN & "</div >"
           XTABLAN = XTablaN & "</div >"

            XTABLAN = XTablaN & chr(13)





        next

           sHTML = sHTML  &   XTABLAN


    end if 








  '  sHTML = sHTML  &  "</DIV >"




    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)


  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


    sHTML = sHTML & "</body></html>"





'''''''''''''

end if

response.write sHTML





%>
