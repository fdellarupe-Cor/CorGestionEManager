<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0


''''''''''''''


 ' XXXX = NEWEMPRE("NO")

 '   volver = OFV.vueltaasp2("../")
 '   volver2 = OFV.vueltaasp2("")

    sNombre = "Tema - Subtema"



 WSPDIVN = "DTTEMAINT"       ' ID DEL DIV DE VENTANA PARA MOSTRAR

 sHTML = ""
 sHTML = sHTML & "<HTML>"
 sHTML = sHTML & "<BODY onload=""window.top.LDVENTANA('" & WSPDIVN & "');"" >"



    Set cn = OFV.conectar(OFV.strconn2)
    
  Fila    = ""



  xTrColor = "#F0F0FF"

  stitem = "Tema"
  sIcono = "ICorTema.jpg"
  sNuevoTipo = "SUBTEMAS"
  nNivel = 0

  sID = 0
  nv = 0
  cborderB = Session("CBRDBOT")
  cborderl = Session("CBRDLFT")
  borderB = ""
  borderl = ""



               sVerEstado2 = "  and Objetos.objnivel >= " & Session("SegNivel") & " "

              if SESSION("EMPMUN") = "S" THEN

                if SESSION("GLBMUN") = 1 THEN
                   sverestado2 = sverestado2 & " and Objetos.OBJUN = " & Session("UN") & " "
                ELSE
                   sverestado2 = sverestado2 & " and "
                   sverestado2 = sverestado2 & " ( (OBJETOS.OBJUN = " & Session("UN") & ") OR "
                   sverestado2 = sverestado2 & "   ( OBJETOS.OBJUN = 1 AND "
                   sverestado2 = sverestado2 & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
                   sverestado2 = sverestado2 & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
                   sverestado2 = sverestado2 & " ) "
                END IF

              end if


                sverestado2 = sverestado2 & " and Objprty.Estado = 3 "

                        xStrSqla = ""
                        xStrSqla  = xStrSqla & " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' "

'''''''''''''''''''''''''''''''''''''''''''
                      if  Session("DELLOG") = "S" then

                          XMFTGRP = split(session("SegGrpLst"),chr(9))
                          xcount  = ubound(XMFTGRP) + 1
                          XMFSEP  = "" 
                          xStrSqla = xStrSqla & " or grpusr in ("

                          for XMFI = 0 to ubound(XMFTGRP)
                              xStrSqla = xStrSqla  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
                              XMFSEP  = "," 
                          next

                          xStrSqla = xStrSqla & ") "

                      else
                          xStrSqla = xStrSqla & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
                      end if 

''''''''''''''''''''''''''''''''''''''''''
                        xStrSqla  = xStrSqla & ")"





  Strsql = "Select * From Tema where aplic = '" & wkaplic & "' or aplic = 'GLB' order by inda, descri"
 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  rs.eof then

      Celda   = ""   
      Celda   = Celda & ofv.GenCelda("","ttt","","2%","","","","&nbsp;","si")
      Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","","No hay elementos asociados","si")

      Fila    = fila & ofv.GenRow("","","","","25","","",Celda,"si")

  else

      Celda   = ""   


           sLink   = "<button value='ver' onclick='chequear();' >"
           Celda   = Celda & ofv.GenCelda("","ttt colspan=2 ","","2%","","","",sLink,"si")
      '     Fila    = ofv.GenRow("","aaa ","","","20","","",Celda,"si")
     


    do until rs.eof

       svatema = "N"

       WTVER = 1

       Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs(0) & " AND ATRMTIPOST = 'T' AND ATRMCODES = " & session("UN")
    ''  response.write strsql & "<br>"
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if  NOT rs1.eof then
           if not RS1(4) then  WTVER = 0
         ' response.write strsql & " " & WTVER & "<br>"
       end if   

       if  WTVER = 1 then
       
           Celda   = ""

         if rs(3) = "R" then

            svatema = "S1"

           sHref   = rs(4) & " "

           if instr(ucase(sHref),"HTTP://") = 0 then sHref   = "HTTP://" & sHref


          ' sHref   = "'" & sHref & "'"


          '''' sImage  = "<font face='corflaticon' size=2 color='#990000'>4</font>"

           sImage  = "<font face='verdana' size=1 color='#990000' style='border-style:solid;border-width:1;border-color:#888888;' ><b>&#149;</b></font>"
           sLink   = ofv.GenLink("","nt",sHref,"_blank","","",sImage,"si")
           Celda   = Celda & ofv.GenCelda("","vt  ","center","2%","","","",sLink,"si")

           sLink   = ofv.GenLink("","nt style='color:#000000;font-weight:bold;font-size:8pt;' ",sHref,"_blank","","",Capitalize(rs(1)),"si")
           Celda   = Celda & ofv.GenCelda("","ut bgcolor='" & xTrColor & "'  ","","90%","","","",sLink,"si")



         else

            svino = "N"
            Fila2 = ""
            xTrColor2 = "#FFFFFF" 

            StrSql  = "Select Prodserv.codprodserv,Prodserv.descri,Prodserv.marcaps,Subtema.descri,"
            StrSql  = StrSql & "Subtema.archivo,Prodserv.segaceso,Prodserv.segactual, Prodserv.indb,Prodserv.PSTIPO ,Prodserv.PSURL "
            StrSql  = StrSql & " From Prodserv, Subtema where "
            StrSql  = StrSql & "Prodserv.marcaps = Subtema.tipobjeto and codtema = " & rs(0)
            StrSql  = StrSql & " order by Prodserv.indb, Prodserv.descri "

            Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
            If not rs2.EOF Then

               Do Until rs2.EOF

                   wpstipo = rs2(8)

                    If  rs2(5) <> "0" Then
                        subtprv = "S"          
                        StrSql  = "Select * From subtemaskr "
                        StrSql  = StrSql & "Where codprodserv = " & rs2(0)
                        StrSql  = StrSql & xStrSqla


                    ''response.write strsql & "<br>"
                        Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
                        If  Not rsIcono.EOF Then
                            subtprv = "N"
                        End If
                        Call OFV.cerrarconsulta(rsIcono)
                    Else
                        subtprv = "N"
                    End If

                    sIcono   = rs2(4)
                    stitem   = rs2(2)
                   ''' sitem    = rs2(3) & ":" & rs2(1)
                    sitem    = Capitalize(rs2(1))
                  '  sitem    = replace(sitem & ""," ","&nbsp;")
                    sitem2   = ofv.ConvertirCar(rs2(3) & ": " & rs2(1)," ","_")
                    stestado = "N"
                    xsf      = ""
                    xsff     = ""



                If subtprv <> "S" Then
                             
                           
                   smas = "N"

                   StrSql = "Select relaobj.codojasoc, objetos.objun  from relaobj, OBJPRTY, Estados, objetos "

                   StrSql = StrSql & " where "

                   StrSql = StrSql & " relaobj.codojasoc = OBJPRTY.codobjeto and "
                   StrSql = StrSql & " objetos.codobjeto = relaobj.codojasoc and "
                   StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
                   StrSql = StrSql & "  objetos.objun =  1 ) and "
                   StrSql = StrSql & " OBJPRTY.estado = Estados.objid and "
                   StrSql = StrSql & " relaobj.tipobjmaster = '" & rs2(2) & "'"
                   StrSql = StrSql & " and relaobj.codobjmaster = " & rs2(0)
                   StrSql = StrSql & " and OBJPRTY.estado = 3 "
               ''  response.write "<br><br><br><br>" & strsql & "<br>"


       Strsql = "select a.cobj, B.objun"
       Strsql = Strsql & " from csrelpath a "
       Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto) "
       Strsql = Strsql & " inner join objprty c on (a.cobj = c.codobjeto) "
       Strsql = Strsql & " inner join estados e on (c.estado = e.objid) "
       Strsql = Strsql & " inner join tipobjeto f on (b.tipobjeto = f.tipobjeto) "

       StrSql = StrSql & " where "
       StrSql = StrSql & " (b.objun =  " & session("UN") & " or "
       StrSql = StrSql & "  b.objun =  1 )  "
       StrSql = StrSql & " and a.stema = " & rs2(0) & " "
       StrSql = StrSql & " and c.estado = 3 " 
       StrSql = StrSql & " and f.TIPOCPT = 'D' " 
       StrSql = StrSql & " and f.filtro = 'S' " 
       StrSql = StrSql & " and (b.aplic = 'GLB' or b.aplic = '" & wkaplic & "') " 

       StrSql = StrSql & " order by b.tipobjeto, b.indobj,  b.descri, e.estado "


                   Set rsmas = OFV.crearconsultaEx(StrSql, cn, 1, "")
                   If  Not rsmas.EOF Then
                       do until rsmas.eof
                          if  cstr(rsmas(1) & "") = cstr(session("UN") & "") then
                              smas = "S"
                              exit do
                          end if 

                       if SESSION("EMPMUN") = "S" THEN               
                          StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
                          StrSql = StrSql & " where "
                          StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rsmas(0) & " and "
                          StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
                          Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
                          If  Not rsmas2.EOF Then
                              smas = "S"
                              exit do
                          end if 
                       end if

                       rsmas.movenext
                       loop   

                   End If
                   Call OFV.cerrarconsulta(rsmas)


                   if smas = "S" then

                      svino = "S"
                      svatema = "S2"

                     if  wpstipo = "R" then
          
                        sHref   = rs2(9) & " "
                        if instr(ucase(sHref),"HTTP://") = 0 then sHref   = "HTTP://" & sHref
                        sLink2   = ofv.GenLink("","cmb style='color:#000000;font-weight:normal;font-size:8pt;' ",sHref,"_BLANK","","",sitem,"si")


                     else

                        if  wspviewopc = 1 then 
                            sxlink  = "Consdocumentos.asp?"
                        else
                            sxlink  = "Consdocumentosweb.asp?"
                        end if

                        sxlink  = sxlink & "ID=" & rs2(0) & "&sitem=" & sitem2 & "&stitem=" & stitem
                        sxlink  = sxlink & "&stnivel=1"

                        sHref   = sxlink
                        sLink2   = ofv.GenLink("","cmb style='color:#000000;font-weight:normal;font-size:8pt;' ",sHref,"","","",sitem,"si")

                     end if 


                     Celda   = ""

                     Celda   = Celda & ofv.GenCelda("","ut style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","","","","",sLink2,"si")
                     Fila2    = Fila2 & ofv.GenRow("","aaa    ","","","","","",Celda,"si")

                     if  xTrColor2 = "#FFFFFF" then
                         xTrColor2 = "#eeeeff" 
                     ELSE
                         xTrColor2 = "#FFFFFF" 
                     END IF

                   end if




                End If

                rs2.movenext
              Loop
            End If


           Celda   = ""


       '''    sImage  = "<font face='corflaticon' size=2 color='#f7aa22'>4</font>"

           sImage1  = "<font id='IMGE" & rs(0) & "' face='verdana' size=1 color='#f7aa22' style='display:block;border-style:solid;border-width:1;border-color:#888888;' ><b>+</b></font>"
           sImage2  = "<font id='IMGM" & rs(0) & "' face='verdana' size=1 color='#f7aa22' style='display:none;border-style:solid;border-width:1;border-color:#888888;' ><b>-</b></font>"
           sLink   = sImage1 & sImage2
           Celda   = Celda & ofv.GenCelda("","vt rowspan=2 valign='top'   ","center","2%","","","",sLink,"si")

           sLink   = Capitalize(rs(1))
           Celda   = Celda & ofv.GenCelda("","ut  bgcolor='" & xTrColor & "' style='border-style:solid;border-width:1;border-color:#dddddd;color:#000000;font-weight:bold;font-size:8pt;' ","","90%","","","",sLink,"si")

           sTABLA2 = ofv.GenTabla("","aaa  ","","100% ","","0","","-2","2","","",Fila2,"si")

           Celda2   = ofv.GenCelda("","ut valign='top' style='border-bottom-style:solid;border-bottom-width:3;border-bottom-color:#ffffff;' ","","90%","","","",stabla2,"si")

         end if

         if  svatema <> "N" then
           if  svatema = "S1" then
               Fila    = Fila & ofv.GenRow("","aaa  ","","","25","","",Celda,"si")
           elseif  svatema = "S2" then
               Fila    = Fila & ofv.GenRow("","aaa  ","","","25","","",Celda,"si")
               Fila    = Fila & ofv.GenRow("","aaa  bgcolor='#fbfbfb' ","","","25","","",Celda2,"si")
           end if


           if  xTrColor = "#F0F0FF" then
               xTrColor = "#F2F2F2" 
           ELSE
               xTrColor = "#F0F0FF" 
           END IF
         end if

           nv      = 1

       end if

      rs.movenext
    loop

  end if


response.write ofv.GenTabla("","aaa  ","","99% ","","0","","0","0","","",Fila,"si")



    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)


 




'''''''''''''

end if

response.write sHTML





%>
