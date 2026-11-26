<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
      dim AposLevel(8)
 		XENTRO1=TRUE
 		XENTRO2=TRUE
 		XENTRO3=TRUE
 		XENTRO4=TRUE
 		XENTRO5=TRUE
 		XENTRO6=TRUE
     ' for i = 0 to 8
     ' AposLevel(i)="A"
     ' next
     ' Amaxlevel="0,0,"
      Generar = ""
      NIVELCONT=90000
      'response.buffer=false
      Set cn = OFV.conectar(OFV.strconn2)
      Set cn2 = OFV.conectar(OFV.strconn0)
      Set cn3 = OFV.conectar(OFV.strconn0)


      strsql="select *  from listadepuestos where tipo='G'"
     ' strsql="select puestodepen from relapuesto where puestodepen>0 group by  puestodepen having  count(*)  >1"
      'response.write strsql
      Set rsasist = OFV.crearconsultaEx(strsql, cn, 1, "")

        if not rsasist.eof then
        objiAsis=rsasist(0)
        else
        objiAsis=-1
        end if
        call ofv.cerrarconsulta(rsasist)

      if request.querystring("mseg")<>"" then
	    session("mseg")=request.querystring("mseg")
      end if

      nID = ""

      sSector = ""
      nSeg = ""

      nID = Request.querystring("ID")

      IF SESSION("MSEG") = "E" THEN

            nidu  = session("SGU_idcostcenter") 'rs(0)
            nivel = session("SGU_idnivel") 'rs(1)

            xxx=getinfouser(session("usuario"))

            If nID = "" then
               nID = nidu
            end if

      ELSE
            nivel=0

            sSector = Request.querystring("Sector")
            sSector = OFV.ConvertirCar(sSector, "_", " ")
            sCodin = "&nbsp;"
            If nID = "" And Request.Form.Count > 0 Then
               nID = Request.Form(1)
            End If
      END IF

      nid = "0" & REPLACE(nid,"nsnc","")

'response.write nid & "---"
      if not nid > "0" then
         sHTML = OFV.MenuHeader("","")
         shtml = shtml & "<br><br><br><br><br><center><p><b><font face=verdana size=4 color=#993333>DEBE SELECCIONAR UN AREA O SECTOR<br>PARA GRAFICAR.</FONT></P>" & chr(13)
         sHTML = sHTML & "<br></body></html>"
      else
         sSECSITE  = 1
         sSECNIVEL = 0
         SNIVELA   = "S/D"
         ndotacion=1
         nivelj0=0
         If nID <> "" Then
            sSQL = "Select descri,codigo,nivelj,dotacion,continuador From listadepuestos Where objid = " & nID
            Set rs = OFV.crearconsultaEx(sSQL, cn, 1, "")
            If Not rs.EOF Then
               sSector   = rs(0)
               scontinuador=rs("continuador")
               ndotacion= rs("dotacion")
               sCodin    = rs(1)
               sSECSITE  = 1
               sSECNIVEL = Request.Form("SECNIVEL")
               nivelj0=rs(2)
               IF  sSECSITE <> "" THEN
                   SESSION("SSITE") = sSECSITE
               ELSE
                   sSECSITE = SESSION("SSITE")
               END IF

               IF sSECNIVEL  = "" THEN sSECNIVEL  = 0



                  sSUCA = "1"
                  sSUCN = "1"



               sSQL = "SELECT * FROM NIVELPUESTO where NIVELP = 0" & RS(2)
               Set rs1 = OFV.crearconsulta2(sSQL, cn)
               If NOT rs1.EOF Then
                  SNIVELA      = RS1(1)
               '''  response.write "NIV " & SNIVELA
               END IF

             End If
             Call OFV.cerrarconsulta(rs)
           End If


        ' response.write NID & "----" & sSector
         sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=2>&#0128;</FONT></B>"
         sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=2>4</FONT></B>"
         sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"

         If nID <> "" Then
            sHTML = OFV.MenuHeader("","")


            shtml=shtml &  chr(13) & "<head><script type='text/javascript' src='incfile/ECOTree.js' ></script><link type='text/css' rel='stylesheet' href='incfile/ECOTree.css' />"
            shtml=shtml &  chr(13) & "<xml:namespace ns='urn:schemas-microsoft-com:vml' prefix='v' />"
            shtml=shtml &  chr(13) & "<style> v\:*{ behavior:url(#default#VML);} </style>"
            shtml=shtml &  chr(13) & "<style>"
            shtml=shtml &  chr(13) & " .copy {"
            shtml=shtml &  chr(13) & " font-family : 'Verdana';"
            shtml=shtml &  chr(13) & " font-size : 9px;"
            shtml=shtml &  chr(13) & " color : #000000;"
            shtml=shtml &  chr(13) & " } "
            shtml=shtml &  chr(13) & "</style>"
            shtml=shtml &  "</head>"



            xin = 0
            ReDim Preserve vecsec(xin)
            ReDim Preserve vecsecd(xin)
            ReDim Preserve vecsecn(xin)
            ReDim Preserve vecALT(xin)

           'Sector Principal

            wwwsettree = ""
         ''   wwwsettreeAN = "400"
         ''   wwwsettreeAL = "200"


            wwwsettreeAN = "290"
            wwwsettreeAL = "150"





            sSQL = "SELECT L_1.objid,L_1.descri,L_1.codigo,L_1.tipo,L_1.segmento,L_1.nivelj,L_1.nivelj,'1','1',LR.puestoREL "
            sSQL = sSQL & " FROM listadepuestos L , RelaPuesto LR, listadepuestos L_1 "
            sSQL = sSQL & " where L.objid = LR.puestoDEPEN and lR.puestoMASTER = L_1.objid and L.objid = " & nID

           response.write " 1   " & SSQL & "<br><br>" & chr(13)
            Set rs = OFV.crearconsulta2(sSQL, cn)
            If rs.EOF Then
               vecsec(xin)  = 0
               vecsecd(xin) = sSector
               vecsecn(xin) = 0
               vecALT(xin)=0
              'nivelj0=0
            Else

                 SCODREL = RS(9)

                 vecsec(xin)  = CLng(rs(0))
                 vecsecd(xin) = sSector
                 vecsecn(xin) = rs(5)
                 vecALT(xin)=1
'''''''''''''
              '  nivelj0=rs(5)

           End If

''response.write session("mseg") & "-" & nid & "-" & nidu & "-" & (cstr(nid) = cstr(nidu))

           if session("mseg") = "E" and cstr(nid) = cstr(nidu) then  vecsec(xin) = 0



           If vecsec(0) <> 0 Then
               ' vecsec(0)=vecsec(0) 'buscapadre(vecsec(0))
                shref = "OrganigramaPuestosgraf.asp?ID=" & vecsec(0) & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")

                sinput = "<input type=button value='Subir' style='color:#000000;font-weight:bold;background-color:#ccddee;' onclick='document.OPSUBE.submit();'  ><form id='OPSUBE' name='OPSUBE' action='" & shref & "' method='post' style='display:none;' >"
                slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")


                wwwsettree = wwwsettree & "myTree.add2(0,-1,'SUBIR',100,60,'#eeeeee','#000000','" & shref & "',null,'1','NO');"


                shrefold = "OrganigramaPuestosgraf.asp?ID=" & vecsec(0) & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")


             '   shref = "OrganigramaPuestosgraf.asp?ID=" & nid & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")

            '    shref = "organigramaUNFunc.asp?ID=" & nid & "&Sector=" & OFV.ConvertirCar(vecsecd(0), " ", "_")


                shref = "GENPUESTOFUNCION.asp?PUESTO=" & nID & "&NOMSECTOR=" & OFV.ConvertirCar(vecsecd(0), " ", "_")
                sinput = "<font face=verdana color=#000080 size=1><b>" & vecsecd(0) & "</b></font>"
                slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

                wwwsettree = wwwsettree & "myTree.add2(1,0,'" &  (vecsecd(0)) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&ndotacion&"','" & scontinuador &"','0');"


           Else


                shref = "GENPUESTOFUNCION.asp?PUESTO=" & nID & "&NOMSECTOR=" & OFV.ConvertirCar(vecsecd(0), " ", "_")
                sinput = "<font face=verdana color=#993333 size=1><b>" & vecsecd(0) & "</b></font>"
                slink = OFV.GenLink("", "nt", shref, "", "", "", sinput, "si")

                wwwsettree = wwwsettree & "myTree.add2(1,-1,'" &  (vecsecd(0)) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'1','NO','0');"

           End If




          sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
          sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, NJ.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR,ER.POSY "
          sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
          sSQL = sSQL & " where E.objid = ER.puestoMASTER and "
          sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
          sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
          sSQL = sSQL & " eco_1.nivelj = NJ.NIVELP  AND "
          sSQL = sSQL & " E.objid = " & nID & "   "
          sSQL = sSQL & " order by ECO_1.nivelj, ECO_1.SECUEN "



          response.write " 2   " & SSQL & "<br><br>" & chr(13)
           Set rs = OFV.crearconsulta2(sSQL, cn)

           'Texto de Sectores Continuadores y Lineas
            Do Until rs.EOF
''''''''''''''''
                 SCODREL = RS(9)

					 xin = xin + 1
					 ReDim Preserve vecsec(xin)
					 ReDim Preserve vecsecd(xin)
					 ReDim Preserve vecsecn(xin)
					 ReDim Preserve vecALT(xin)
					 vecsec(xin) = CLng(rs(0))
					 vecsecd(xin) = rs(1)
					 vecsecn(xin) = rs(5)
                     vecALT(xin)=2
                     nivelj1=rs(5)
                 xsegmentoSALTAR=CLNG("0" & RS("POSY"))

                 xsegmento = "0" & rs(4)

                 xgrafica = "" & rs(13)

					 if xgrafica = "S" then
						shref = "OrganigramaPuestosgraf.asp?ID=" & vecsec(xin) & "&Sector=" & OFV.ConvertirCar(vecsecd(xin), " ", "_")
					 else
						shref = "GENPUESTOFUNCION.asp?PUESTO=" & vecsec(xin) & "&NOMSECTOR=" & OFV.ConvertirCar(vecsecd(xin), " ", "_")
					 end if

                     pad=1
				    xsegmentoSALTAR=(nivelj1-nivelj0)- 1
				    'response.write nivelj1 & "---" & nivelj0 & "<br>"
                     IF xsegmentoSALTAR > 0 THEN
                      pad=AGREGARNIVEL(xsegmentoSALTAR,1)
                 '     vecALT(xin)=vecALT(xin) +  xsegmentoSALTAR

                     END IF

					 wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  rs(1) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&RS("DOTACION")&"','"&RS("CONTINUADOR")&"','"&rs(0)&"');"

                 xinw = xin + 1

                ' IF AposLevel(2) < RS("NIVELP") THEN
                '  AposLevel(2)= RS("NIVELP")
                ' END IF
'''''''''''''''''''''''

        '  sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
        '  sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN "
        '  sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento "
        '  sSQL = sSQL & " where E.objid = ER.puestoMASTER and "
        '  sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
        '  sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
        '  sSQL = sSQL & " E.objid = " & rs(0) & "   "
        '  sSQL = sSQL & " and segmento.codseg =  '" & xsegmento & "' "
        '  sSQL = sSQL & " order by ECO_1.nivelj, ECO_1.SECUEN "


          sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
          sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, NJ.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR,ER.POSY  "
          sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
          sSQL = sSQL & " where E.objid = ER.puestoMASTER and "
          sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
          sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
          sSQL = sSQL & " eco_1.nivelj = NJ.NIVELP  AND "
          sSQL = sSQL & " E.objid = " & rs(0) & "   "
          sSQL = sSQL & " and (segmento.codseg =  '" & xsegmento & "' OR ECO_1.TIPO='GA') "
          sSQL = sSQL & " order by ECO_1.nivelj, ECO_1.SECUEN "


           response.write " 3   " & SSQL & "<br><br>" & chr(13)
           Set rsn = OFV.crearconsulta2(sSQL, cn)

           'Texto de Sectores Continuadores y Lineas
            Do Until rsn.EOF
''''''''''''''''
                 xSCODREL = RSn(9)

                 xin = xin + 1

                 xvecsec = CLng(rsn(0))
                 xvecsecd = rsn(1)
                 xvecsecn = rsn(5)
                 xsegmentoSALTAR=CLNG("0" & RSN("POSY"))
                 xsegmento2 = "0" & rsn(4)
                 nivelj2=rsn(5)
                 ReDim Preserve vecALT(xin)

                 xgrafica = "" & rsn(13)

                 if xgrafica = "S" then
                    shref = "OrganigramaPuestosgraf.asp?ID=" & xvecsec & "&Sector=" & OFV.ConvertirCar(xvecsecd, " ", "_")
                 else
                    shref = "GENPUESTOFUNCION.asp?PUESTO=" & xvecsec & "&NOMSECTOR=" & OFV.ConvertirCar(xvecsecd, " ", "_")
                 end if


             			 pad=xinw
                         xsegmentoSALTAR=(nivelj2-nivelj1 )- 1
			             IF xsegmentoSALTAR > 0 THEN
			              pad=AGREGARNIVEL(xsegmentoSALTAR,xinw)
			          '    vecALT(xin)=vecALT(xin) +  xsegmentoSALTAR

                         END IF

                ' wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&RSN("DOTACION")&"','"&RSN("CONTINUADOR")&"','"&rsn(0)&"');"
          '        if clng(RSn(0))=objiAsis AND XENTRO1 THEN
		'		 XENTRO1=FALSE
		'		 wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN("DOTACION")&"','"&RSN("CONTINUADOR")&"','"&rsn(0)&"');"
		'	   ELSEIF  clng(RSn(0))=objiAsis  THEN
		 		if clng(RSn(0))=objiAsis  THEN
				wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN("DOTACION")&"','SI','"&rsn(0)&"');"
			   ELSE
				 wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&RSN("DOTACION")&"','"&RSN("CONTINUADOR")&"','"&rsn(0)&"');"

               END IF



                 xinw2 = xin + 1

'''''''''''''''''''''''

       '   sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
       '   sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN "
       '   sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento "
       '   sSQL = sSQL & " where E.objid = ER.puestoMASTER and "
       '   sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
       '   sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
       '   sSQL = sSQL & " E.objid = " & rsn(0) & "   "
       '   sSQL = sSQL & " and segmento.codseg =  '" & xsegmento2 & "' "
       '   sSQL = sSQL & " order by ECO_1.nivelj, ECO_1.SECUEN "


          sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
          sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, NJ.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR,ER.POSY "
          sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
          sSQL = sSQL & " where E.objid = ER.puestoMASTER and "
          sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
          sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
          sSQL = sSQL & " eco_1.nivelj = NJ.NIVELP  AND "
          sSQL = sSQL & " E.objid = " & rsn(0) & "   "
          sSQL = sSQL & " and (segmento.codseg =  '" & xsegmento2 & "' OR ECO_1.TIPO='GA') "
          sSQL = sSQL & " order by ECO_1.nivelj, ECO_1.SECUEN "


            response.write " 4   " & SSQL & "<br><br>" & chr(13)
           Set rsn2 = OFV.crearconsulta2(sSQL, cn)

           'Texto de Sectores Continuadores y Lineas
            Do Until rsn2.EOF
''''''''''''''''
                 xSCODREL2 = RSn2(9)

                 xin = xin + 1

                 xvecsec2 = CLng(rsn2(0))
                 xvecsecd2 = rsn2(1)
                 xvecsecn2 = rsn2(5)
                  xsegmentoSALTAR=CLNG("0" & RSN2("POSY"))
                 xsegmento3 = "0" & rsn2(4)
                 nivelj3=rsn2(5)
                 xgrafica = "" & rsn2(13)

                 if xgrafica = "S" then
                    shref = "OrganigramaPuestosgraf.asp?ID=" & xvecsec2 & "&Sector=" & OFV.ConvertirCar(xvecsecd2, " ", "_")
                 else
                    shref = "GENPUESTOFUNCION.asp?PUESTO=" & xvecsec2 & "&NOMSECTOR=" & OFV.ConvertirCar(xvecsecd2, " ", "_")
                 end if

           		pad=xinw2
                xsegmentoSALTAR=(nivelj3-nivelj2 )- 1
				IF xsegmentoSALTAR > 0 THEN
				 pad=AGREGARNIVEL(xsegmentoSALTAR,xinw2)
  			    END IF
                ' wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd2) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&RSN2("DOTACION")&"','"&RSN2("CONTINUADOR")&"','"&rsn2(0)&"');"
             '    if clng(RSn2(0))=objiAsis AND XENTRO2 THEN
			'	                 XENTRO2=FALSE
			'	                 wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd2) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN2("DOTACION")&"','"&RSN2("CONTINUADOR")&"','"&rsn2(0)&"');"
			'	               ELSEIF  clng(RSn2(0))=objiAsis  THEN
			   if clng(RSn2(0))=objiAsis  THEN
				                wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd2) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN2("DOTACION")&"','SI','"&rsn2(0)&"');"
				               ELSE
				                 wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd2) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&RSN2("DOTACION")&"','"&RSN2("CONTINUADOR")&"','"&rsn2(0)&"');"

               END IF




                 xinw3 = xin + 1

'''''''''''''''''''''''


          sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
          sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, NJ.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR,ER.POSY "
          sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
          sSQL = sSQL & " where E.objid = ER.puestoMASTER and "
          sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
          sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
          sSQL = sSQL & " eco_1.nivelj = NJ.NIVELP  AND "
          sSQL = sSQL & " E.objid =" & rsn2(0) & "   "
          sSQL = sSQL & " and (segmento.codseg =  '" & xsegmento3 & "' OR ECO_1.TIPO='GA') "
          sSQL = sSQL & " order by ECO_1.nivelj, ECO_1.SECUEN "


          response.write " 5   " & SSQL & "<br><br>" & chr(13)
           Set rsn3 = OFV.crearconsulta2(sSQL, cn)

           'Texto de Sectores Continuadores y Lineas
            Do Until rsn3.EOF
''''''''''''''''
                 xSCODREL3 = RSn3(9)

                 xin = xin + 1

                 xvecsec3 = CLng(rsn3(0))
                 xvecsecd3 = rsn3(1)
                 xvecsecn3 = rsn3(5)
                  xsegmentoSALTAR=CLNG("0" & RSN3("POSY"))
                  xsegmento4 = "0" & rsn3(4)
                  nivelj4=rsn3(5)
                 xgrafica = "" & rsn3(13)

                 if xgrafica = "S" then
                    shref = "OrganigramaPuestosgraf.asp?ID=" & xvecsec3 & "&Sector=" & OFV.ConvertirCar(xvecsecd3, " ", "_")
                 else
                    shref = "GENPUESTOFUNCION.asp?PUESTO=" & xvecsec3 & "&NOMSECTOR=" & OFV.ConvertirCar(xvecsecd3, " ", "_")
                 end if

                 pad=xinw3
                xsegmentoSALTAR=(nivelj4-nivelj3 )- 1
			    iF xsegmentoSALTAR > 0 THEN
			 	 pad=AGREGARNIVEL(xsegmentoSALTAR,xinw3)
  			    END IF

              ' if clng(RSn3(0))=objiAsis AND XENTRO3 THEN
             '    XENTRO3=FALSE
             '    wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd3) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN3("DOTACION")&"','"&RSN3("CONTINUADOR")&"','"&rsn3(0)&"');"
             '  ELSEIF  clng(RSn3(0))=objiAsis  THEN
              if clng(RSn3(0))=objiAsis  THEN
                wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd3) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN3("DOTACION")&"','SI','"&rsn3(0)&"');"
               ELSE
                 wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd3) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&RSN3("DOTACION")&"','"&RSN3("CONTINUADOR")&"','"&rsn3(0)&"');"

               END IF



                 ' otro nivel mas

                  sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
				           sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, NJ.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR,ER.POSY "
				           sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
				           sSQL = sSQL & " where E.objid = ER.puestoMASTER and "
				           sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
				           sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
				           sSQL = sSQL & " eco_1.nivelj = NJ.NIVELP  AND "
				           sSQL = sSQL & " E.objid =" & rsn3(0) & "   "
				           sSQL = sSQL & " and (segmento.codseg =  '" & xsegmento4 & "' OR ECO_1.TIPO='GA') "
				           sSQL = sSQL & " order by ECO_1.nivelj, ECO_1.SECUEN "
				            xinw4 = xin + 1

				        response.write " 5   " & SSQL & "<br><br>" & chr(13)
				            Set rsn4 = OFV.crearconsulta2(sSQL, cn)

				            'Texto de Sectores Continuadores y Lineas
				             Do Until rsn4.EOF
				 ''''''''''''''''
				                  xSCODREL4 = RSn4(9)

				                  xin = xin + 1

				                  xvecsec4 = CLng(rsn4(0))
				                  xvecsecd4 = rsn4(1)
				                  xvecsecn4 = rsn4(5)
                                  xsegmentoSALTAR=CLNG("0" & RSN4("POSY"))
				                  xgrafica = "" & rsn4(13)
                                   xsegmento5 = "0" & rsn4(4)
                                   nivelj5=rsn4(5)
				                  if xgrafica = "S" then
				                     shref = "OrganigramaPuestosgraf.asp?ID=" & xvecsec4 & "&Sector=" & OFV.ConvertirCar(xvecsecd4, " ", "_")
				                  else
				                     shref = "GENPUESTOFUNCION.asp?PUESTO=" & xvecsec4 & "&NOMSECTOR=" & OFV.ConvertirCar(xvecsecd4, " ", "_")
				                  end if

                                pad=xinw4
                                 xsegmentoSALTAR=(nivelj5-nivelj4 )- 1


								IF xsegmentoSALTAR > 0 THEN
								  pad=AGREGARNIVEL(xsegmentoSALTAR,xinw4)
  			   					END IF
				           '     if clng(RSn4(0))=objiAsis AND XENTRO4 THEN
				           '       XENTRO4=FALSE
				           '       wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd4) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN4("DOTACION")&"','"&RSN4("CONTINUADOR")&"','"&rsn4(0)&"');"
				           '     ELSEIF  clng(RSn4(0))=objiAsis  THEN
				            if clng(RSn4(0))=objiAsis  THEN
				                 wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd4) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN4("DOTACION")&"','SI','"&rsn4(0)&"');"
				                ELSE
				                  wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd4) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&RSN4("DOTACION")&"','"&RSN4("CONTINUADOR")&"','"&rsn4(0)&"');"

				                END IF



				                  ' otro nivel mas

				           ' otro nivel mas

                  sSQL = "SELECT DISTINCT ECO_1.objid,ECO_1.descri,ECO_1.codigo,Eco_1.tipo,Eco_1.segmento,Eco_1.nivelj,Eco_1.nivelj,"
				           sSQL = sSQL & " '1','1',ER.puestorel,ER.puestoMASTER,ER.puestoDEPEN, ECO_1.SECUEN, NJ.grafica,ECO_1.DOTACION,ECO_1.CONTINUADOR,ER.POSY "
				           sSQL = sSQL & " FROM listadepuestos E, relapuesto ER, listadepuestos  ECO_1, segmento, NIVELPUESTO NJ "
				           sSQL = sSQL & " where E.objid = ER.puestoMASTER and "
				           sSQL = sSQL & " ER.puestoDEPEN = ECO_1.objid and "
				           sSQL = sSQL & " eco_1.segmento = segmento.codseg  AND "
				           sSQL = sSQL & " eco_1.nivelj = NJ.NIVELP  AND "
				           sSQL = sSQL & " E.objid =" & rsn4(0) & "   "
				           sSQL = sSQL & " and (segmento.codseg =  '" & xsegmento5 & "' OR ECO_1.TIPO='GA') "
				           sSQL = sSQL & " order by ECO_1.nivelj, ECO_1.SECUEN "
				            xinw5 = xin + 1

				         response.write " 7   " & SSQL & "<br><br>" & chr(13)
				            Set rsn5 = OFV.crearconsulta2(sSQL, cn)

				            'Texto de Sectores Continuadores y Lineas
				             Do Until rsn5.EOF
				 ''''''''''''''''
				                  xSCODREL5 = RSn5(9)

				                  xin = xin + 1

				                  xvecsec5 = CLng(rsn5(0))
				                  xvecsecd5 = rsn5(1)
				                  xvecsecn5 = rsn5(5)
                                  xsegmentoSALTAR=CLNG("0" & RSN5("POSY"))
                                  nivelj6=rsn5(5)
				                  xgrafica = "" & rsn5(13)

				                  if xgrafica = "S" then
				                     shref = "OrganigramaPuestosgraf.asp?ID=" & xvecsec5 & "&Sector=" & OFV.ConvertirCar(xvecsecd5, " ", "_")
				                  else
				                     shref = "GENPUESTOFUNCION.asp?PUESTO=" & xvecsec5 & "&NOMSECTOR=" & OFV.ConvertirCar(xvecsecd5, " ", "_")
				                  end if

                                pad=xinw5
                                 xsegmentoSALTAR=(nivelj6-nivelj5 )- 1
								IF xsegmentoSALTAR > 0 THEN
								  pad=AGREGARNIVEL(xsegmentoSALTAR,xinw5)
  			   					END IF
				      '          if clng(RSn5(0))=objiAsis AND XENTRO5 THEN
				      '            XENTRO5=FALSE
				      '            wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd5) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN5("DOTACION")&"','"&RSN5("CONTINUADOR")&"','"&rsn5(0)&"');"
				      '          ELSEIF  clng(RSn5(0))=objiAsis  THEN
				       			if clng(RSn5(0))=objiAsis  THEN
				                 wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd5) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "','"&objiAsis&"','"&RSN5("DOTACION")&"','SI','"&rsn5(0)&"');"
				                ELSE
				                  wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & "," & pad & ",'" &  (xvecsecd5) & "'," & wwwsettreeAN & "," & wwwsettreeAL & ",'','#000000','" & shref & "',null,'"&RSN5("DOTACION")&"','"&RSN5("CONTINUADOR")&"','"&rsn5(0)&"');"

				                END IF



				                  ' otro nivel mas


				             rsn5.movenext
            Loop


				             rsn4.movenext
            Loop


            rsn3.movenext
            Loop



''''''''''''''''''''



            rsn2.movenext
            Loop



''''''''''''''''''''


            rsn.movenext
            Loop



''''''''''''''''''''

            rs.movenext
            Loop
          Call OFV.cerrarconsulta(rs)



            xinS = 0
            ReDim Preserve vecsecS(xinS)
            ReDim Preserve vecsecSD(xinS)
            ReDim Preserve vecsecSN(xinS)



               sft  = "<font face=verdana color=#000077 size=1>&nbsp;"
	       sft2 = "<font face=verdana color=#770000 size=1><b>&nbsp;"
	       sft3 = "<font face=verdana color=#ff0000 size=1>&nbsp;"
               sft1 = "<font face=verdana color=#770000 size=1>&nbsp;"
               sff = "</font>"






            shtml=shtml & " <div id='myTreeContainer' style='position:absolute;top:150px;left:1px' > "
            shtml=shtml & " </div> "


            shtml=shtml &  "<script>"
            shtml=shtml &  " var myTree = null;"
            delim=""
            Amaxlevel=""
          '  for i=0 to 8
          '  select case AposLevel(i)
          '  case  "A"
          '   Amaxlevel=Amaxlevel & DELIM & "10"
          '  case  "B"
          '   Amaxlevel=Amaxlevel & DELIM & "110"
          '  case  "C"
          '   Amaxlevel=Amaxlevel & DELIM & "210"
          '  case  "D"
          '   Amaxlevel=Amaxlevel & DELIM & "310"
          '  END SELECT
          '  DELIM=","
          '  next

           ' RESPONSE.WRITE Amaxlevel
            shtml=shtml &  "var Xasis=true;"


            shtml=shtml &  " function CreateTree() { "
            shtml=shtml &  "myTree = new ECOTree('myTree','myTreeContainer');"
        '''    shtml=shtml &  "myTree.config.colorStyle = ECOTree.CS_LEVEL;"
            shtml=shtml &  "myTree.config.colorStyle = ECOTree.CS_NODE;"
            shtml=shtml &  "myTree.config.nodeFill = ECOTree.NF_GRADIENT;"
            shtml=shtml &  "myTree.config.useTarget = true;"
            shtml=shtml &  "myTree.config.selectMode = ECOTree.SL_NONE;"
            shtml=shtml &  "myTree.config.defaultNodeWidth = 85;"
            shtml=shtml &  "myTree.config.defaultNodeHeight = 60;"
            shtml=shtml &  "myTree.config.iSubtreeSeparation = 40;"
            shtml=shtml &  "myTree.config.iSiblingSeparation = 40;"
            shtml=shtml &  "myTree.config.iLevelSeparation = 150;"
             shtml=shtml &  "myTree.config.nodeColor = "&chr(34)&"#77AADD"&chr(34)&";"
         '   shtml=shtml &  "myTree.config.iNodeJustification = ECOTree.NJ_CENTER; "
            shtml=shtml &  "myTree.config.linkColor = 'black' ; "
            shtml=shtml &  "myTree.config.iNodeJustification =ECOTree.NJ_CENTER; "

            shtml=shtml &  wwwsettree

            shtml=shtml &  "myTree.UpdateTree();"
            shtml=shtml &  "}"

            shtml=shtml &  " CreateTree(); "

            shtml=shtml &  "</script>"



      Else


      End If

      Call OFV.cerrarconn(cn)
      Call OFV.cerrarconn(cn2)
      Call OFV.cerrarconn(cn3)


      if request.querystring("mseg")="M" then
         sHTML = sHTML  & "</body></html>"
      else
          sHTML = sHTML  & "<br></body></html>"
      end if


     Generar = sHTML

  end if





end if

'for i=0 to 8
'response.write  AposLevel(i)  & "   " &i & "<br>"
'next

function buscapadre(idp)
SPUESTO=idp
SPUESTOx=""
Strsql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador FROM RelaPuesto  "
		 strsql = strsql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid) "
         Strsql = Strsql & " WHERE RelaPuesto.puestodepen = 0" & spuesto & "  ORDER BY RelaPuesto.puestorel "
       '  RESPONSE.WRITE STRSQL
		  set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	      if not rsX.eof then

	      do while not rsX.eof
	       SPUESTOx=RSX("PUESTOMASTER")
	       if rsX("continuador")="NO" THEN EXIT DO
	       call ofv.cerrarconsulta(rsX)
	       Strsql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador FROM RelaPuesto  "
		   strsql = strsql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid) "
	       Strsql = Strsql & " WHERE RelaPuesto.puestodepen = 0" & spuestox & "  ORDER BY RelaPuesto.puestorel "

	       set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	      loop
	      end if
	      if SPUESTOx="" then SPUESTOx=SPUESTO
buscapadre=SPUESTOx
end function


FUNCTION AGREGARNIVEL(NIVEL,PADRE)


	FOR IX=1 TO NIVEL
	xin = xin + 1

	 wwwsettree = wwwsettree & "myTree.add2(" & (xin + 1) & ","& PADRE &",'virtualD'," & wwwsettreeAN & "," & wwwsettreeAL & ",'OrganigramaPuestosgraf.asp?ID=17&Sector=GERENTE_DE_AREA_RIESGO_CREDITICIO','#000000','',null,'0','SI',"&xin&");"

	 PADRE=XIN+1
	 xin = xin + 1
   NEXT

   AGREGARNIVEL=padre
END FUNCTION

Response.Write sHTML


%>