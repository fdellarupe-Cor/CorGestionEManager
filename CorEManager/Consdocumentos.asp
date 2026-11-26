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

    xs = xs & " var mydat= 'Consultar.asp?Codigo=' + c1 + '&Tipo=' + c2 + '&Perfil=' + c3 + '&Datos=&Volver=" & volver2 & "'; "

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



  Fila = ""


   sHTML = sHTML & "<br><center>"


    Response.write sHTML
    sHTML = ""
    Response.flush

    Set cn = OFV.conectar(OFV.strconn2)
    
    xTrColor = "#FFFFFF"     
   

    
    

    If Stnivel = 1 Then

        Session("XstemaSid")     = sID
        Session("XstemaSitem")   = sitem
        Session("XstemaStitem")  = stitem
        Session("XstemaStnivel") = Stnivel
        Session("XstemaStesq")   = stesq

    Else

        sesn                      = "X" & Stnivel
        Session(sesn & "Sid")     = sID
        Session(sesn & "Sitem")   = sitem
        Session(sesn & "Stitem")  = stitem
        Session(sesn & "Stnivel") = Stnivel
        Session(sesn & "Stesq")   = stesq

    End If



        stitemt     = "Tema"
        sIconot     = "ICorTema.jpg"
        sNuevoTipot = "SUBTEMAS"
        nNivelt     = 0

        StrSql = "Select codtema from prodserv where codprodserv = 0" & Session("XstemaSid")
        Set rsx = OFV.crearconsultaEx(StrSql, cn, 1, "")
        If Not rsx.EOF Then
             Strsql = "Select descri, codtema From Tema where codtema = " & rsx(0)
             Set rsx = OFV.crearconsultaEx(StrSql, cn, 1, "")
             If Not rsx.EOF Then
                stemaid   = rsx(1)
                stemades  = rsx(0)
             End If
        End If
        Call OFV.cerrarconsulta(rsx)



        sHref = "ObjTreeSubTemaCons.asp?Tipo=" & sNuevoTipot & "&ID=" & stemaid & "&sitem=TEMA:_"
        sHref = sHref & ofv.convertircar(stemades," ","_") & "&stitem="
        sHref = sHref & ofv.convertircar(stitemt," ","_") & "&sticono=" & sIconot
        sHref = sHref & "&stnivel=" & nNivelt


        sHref2 = "ConsObjetos.asp?ID=" & Session("XstemaSid")
        sHref2 = sHref2 & "&sitem=" & OFV.convertircar(Session("XstemaSitem"), " ", "_") & "&stitem="
        sHref2 = sHref2 & OFV.convertircar(Session("XstemaStitem"), " ", "_")
        sHref2 = sHref2 & "&stnivel=" & Session("XstemaStnivel") & "&stesq=" & Session("XstemaStesq")



      TABLA   = ofv.GenTabla("","aaa  style='border-style=solid;border-width:1;border-color:#aaaaaa;' ","","98% ","","0","","-1","3","","","","no")
      sHTML   = sHTML & Tabla
'''''''''''''''''''''''''''''

''''''''''''''''' TEMA ''''''''''''
      Celda   = ""

      sImage  = "<font face='corflaticon' size=3 color='#f7aa22'>4</font>"
      sLink   = sImage
      Celda   = Celda & ofv.GenCelda("","uts style='border-style=solid;border-width:1;border-color:#aaaaaa;' ","center","2%","","","",sLink,"si")

      sLink   = "&nbsp;<B><font color='#0000AA' size=2 >&#0124;</font></B>&nbsp;" & "TEMA:&nbsp;" & stemades & "&nbsp;<B><font color='#0000AA' size=2 >&#0124;</font></B>&nbsp;" & Session("XstemaSitem") & "&nbsp;<B><font color='#0000AA' size=2 >&#0124;</font></B>&nbsp;"
      Celda   = Celda & ofv.GenCelda("","ut style='color:#000000;' bgcolor='#f0f0f0' ","","95%","","","",sLink,"si")

      Fila    = Fila & ofv.GenRow("","aaa  ","","","25","","",Celda,"si")

      sHTML   = sHTML & Fila
      sHTML = sHTML & "</td>"
      sHTML = sHTML & "</tr>"
      sHTML = sHTML & "</table>"


      sHTML = sHTML & "<br>"
 
      Fila    = ""
      sHTML   = sHTML & "<table cellpadding=3 cellspacing=2  STYLE='border-style:solid;border-width:1;border-color:#ffffff;' >"


       Strsql = "select a.cobj, B.DESCRI, b.tipobjeto, e.estado,"
       Strsql = Strsql & " b.tipobjeto, c.fecemis, c.fecvigen, c.fecbaja, "
       StrSql = StrSql & " c.autor, c.aprobo, c.elimino, c.fecvto, b.OBJUN, b.indobj, "
       Strsql = Strsql & "  f.descri as tobjdescri , f.perfil, f.TIPOCPT  "
       Strsql = Strsql & " from csrelpath a "
       Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto) "
       Strsql = Strsql & " inner join objprty c on (a.cobj = c.codobjeto) "
       Strsql = Strsql & " inner join estados e on (c.estado = e.objid) "
       Strsql = Strsql & " inner join tipobjeto f on (b.tipobjeto = f.tipobjeto) "

       StrSql = StrSql & " where "
       StrSql = StrSql & " (b.objun =  " & session("UN") & " or "
       StrSql = StrSql & "  b.objun =  1 )  "
       StrSql = StrSql & " and a.stema = " & sID & " "
       StrSql = StrSql & " and c.estado = 3 " 
       StrSql = StrSql & " and f.TIPOCPT = 'D' " 

       StrSql = StrSql & " order by b.tipobjeto, b.indobj,  b.descri, e.estado "





  '  StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
  '  StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
  '  StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN, objetos.indobj"
  '  StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
  '  StrSql = StrSql & " where "
  '  StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
  '  StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
  '  StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
  '  StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
  '  StrSql = StrSql & "  objetos.objun =  1 )  and "
  '  StrSql = StrSql & " RELAOBJ.codobjmaster = " & sID
  '  StrSql = StrSql & " and relaobj.tipobjmaster = '" & stitem & "' " & sVerEstado
  '  StrSql = StrSql & " order by objetos.indobj, OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
    
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


        Fila    = ofv.GenRow("","aaa  ","","","25","","","","no")
        sHTML = sHTML & Fila

        sVino = "N"

        nlinea = 0
        Do Until rs.EOF
             sVinook = "N"
             if  cstr(rs(12) & "") = cstr(session("UN") & "") then
   '' response.write "<br><br>entro"
                 sVino   = "S"
                 sVinook = "S"
             end if 

          sNIVE   = "D"
          sitem = sTipoE & ": " & rs(1)

          if  sVinook = "S" then

            stitem = rs(4)

          '  StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
          '  StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs(4) & "'"
          '  Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
          '  If Not rsIcono.EOF Then
          '      sIcono  = rsIcono(0)
          '      sTipoE  = rsIcono(1)
          '      sPerfil = rsIcono(2)
          '      sNIVE   = rsIcono(3)
          '   '''   sitem   = sTipoE & ": " & rs(1) & " - (" & rs(3) & ")"
          '      sitem = Capitalize(rs(1))
          '     ' IF sNIVE = "D" THEN  sitem = sitem & " - (" & rs(3) & ")"
          '  End If
          '  Call OFV.cerrarconsulta(rsIcono)



                sTipoE  = rs(14)
                sPerfil = rs(15)
                sNIVE   = rs(16)

                sitem = Capitalize(rs(1))

            if tipoant <> sTipoE then

               wspleftb = ""

               if tipoant <> "NO" then 


                  sHTML   = sHTML & "</table >"

                  Fila    = ofv.GenFCelda()
                  sHTML = sHTML & Fila

               else

                  wspleftb = "border-left-style:solid;border-left-width:1;border-left-color:#aa0000;"  


               end if

               tipoant = sTipoE

                  canttipo = canttipo + 1

                  if canttipo > 4 then

                     Fila    = ofv.GenFRow()
                     sHTML = sHTML & Fila
                     sHTML   = sHTML & "</table >" & "<br>"

                     sHTML   = sHTML & "<table cellpadding=3 cellspacing=2   STYLE='border-style:solid;border-width:1;border-color:#ffffff;' >"
                     Fila    = ofv.GenRow("","aaa  ","","","25","","","","no")
                     sHTML = sHTML & Fila
                     canttipo = 1

                     wspleftb = "border-left-style:solid;border-left-width:1;border-left-color:#aa0000;"  

                  end if  


                  Fila    = ofv.GenCelda("","ut valign=top width=260 STYLE='" & wspleftb & "border-right-style:solid;border-right-width:1;border-right-color:#aa0000;' ","center","","","","","","no")
                  sHTML = sHTML & Fila

                  sHTML   = sHTML & "<table cellpadding=1 cellspacing=2 width=250   >"

                  Celda   = ""
                  Celda   = Celda & ofv.GenCelda("","vti colspan=2 STYLE='background-color:#f0eedd;color:#000000;font-size:9pt;border-style:solid;border-width:1;border-color:#aaaaaa;' ","","","","","",Capitalize(tipoant),"si")
                  Fila    = ofv.GenRow("","aaa  ","","","25","","",Celda,"si")
                  sHTML = sHTML & Fila

            end if 



            IF sNIVE = "C" then


               Strsql = "Select * From OBJSDK Where tipobjeto='" & rs(4) & "' and codobjeto = " & rs(0)
               set rsIcono = ofv.crearconsultaEx(StrSql,cn,1,parametros)
               if not rsIcono.eof then

      
                  if Session("Usuario") <> rsIcono(7) and rsIcono(5) then


                     call ofv.cerrarconsulta(rsIcono)


                     Strsql = "Select * From OBJSDKDET Where tipobjeto='" & rs(4)
                     Strsql = Strsql & "' and codobjeto = "  & rs(0)
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

'response.write strsql & "<br>"
                      set rsIcono = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                      if rsIcono.eof then sVinook = "N"

                  end if
                end if
               call ofv.cerrarconsulta(rsIcono)



            end if



          sxlink = ""
          sHref   = sxlink
          scon     = "&nbsp;"



            If clng("0" & rs(12)) = 1 and  SESSION("UN") <> 1 and  SESSION("EMPMUN") = "S" Then
                smasc = "&nbsp;&nbsp;&nbsp;<font color=#993333 ><font class=btic title='Corporativo' >G</font></font>"
            Else
                smasc = "&nbsp;"
            End If


            Celda   = ""
           ' Celda   = Celda & ofv.GenCelda("","ut","center","1%","","","",smasxxx,"si")
          '  Celda   = Celda & ofv.GenCelda("","uts","center","2%","","","",scon,"si")


              sImage  = "<font face='corflaticon' size=4 color='#005500' >&#0157;</font>"
              sHref    = "&nbsp;"
              sFuncion = " onclick=" & Chr(34) & " Axn('" & rs(0) & "','" & stitem & "','" & sPerfil & "'); return false;" & Chr(34)

              scon     = "<A class=nt href='" & sHref & "' " & sFuncion & " title='ver...'  >"
              scon     = scon & sImage
              scon     = scon & "</a>"

              sLink    = scon
            '  Celda   = Celda & ofv.GenCelda("","uts","center","2%","","","",sLink,"si")


            '  scon     = "<A class=nt href='" & sHref & "' " & sFuncion & " title='verq...'  >"
            '  scon     = scon & sitem & smasc & "&nbsp;&nbsp;&nbsp;&nbsp" & smas
            '  scon     = scon & "</a>"


             ''''' sLink    = "&nbsp;<font color='#AA0000' size=2 >&#0149;</font>&nbsp;&nbsp;&nbsp;" & sitem
              sLink1    = "&nbsp;<font color='#AA0000' size=2 >&#0149;</font>&nbsp;"
              sLink    = "&nbsp;&nbsp;" & sitem
              Celda   = Celda & ofv.GenCelda("","uts","center","2%","","","",sLink1,"si")
              Celda   = Celda & ofv.GenCelda("","ut " & sFuncion & " title='Consultar.' STYLE='cursor:pointer;background-color:" & xTrColor & ";color:#000000;font-size:9pt;border-style:solid;border-width:1;border-color:#dddddd;' ","","95%","","","",sLink,"si")




            Fila    = ofv.GenRow("","aaa  ","","","25","","",Celda,"si")
            sHTML = sHTML & Fila

            if  xTrColor = "#FFFFFF" then
                xTrColor = "#F0F0F0" 
            ELSE
                xTrColor = "#FFFFFF" 
            END IF

          end if

       '  end if


            rs.movenext
            nlinea = nlinea + 1
        Loop

        Fila    = ofv.GenFRow()
        sHTML = sHTML & Fila


    End If

    if  sVino = "N" then

        Celda   = ""

        sLink   = "&nbsp;"
        Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

        sLink   = "&nbsp;"
        Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

        sLink   = "No hay Esquemas asociados"
        Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","",sLink,"si")

        Fila    = Fila & ofv.GenRow("","aaa bgcolor='" & xTrColor & "' ","","","25","","",Celda,"si")
        sHTML = sHTML & Fila
    end if 


    sHTML = sHTML & "</table>"






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
