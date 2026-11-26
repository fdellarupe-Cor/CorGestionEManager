<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  svision       = request.querystring("VISION")
  svisionMAIN   = request.querystring("VISIONMAIN")
  sACTIVRIESGO  = request.querystring("ACTIVRIESGO")
  sFECHA        = request.querystring("FECHA")

  if svision = "" then svision = "AGN"

  sHTML = ofv.FormHeader(session("FormN"))



  WKVISION = "AGN"

 






       WKVISION = "AGN"
       WKVISDES = "AGENDA DE SEGUIMIENTO"
       sImage = "AGENDA"



    Celda = ""


    
   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","SEGUIMIENTO PLAN DE MITIGACION","si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","","&nbsp;","si")  
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;BORDER-COLOR:#888888;border-width:1;BACKGROUND-COLOR:#E0E0E0;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")




  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



  DIM TABRANRIESGOS(10,5) 


  Strsql = "Select * From ORRANGORIESGO ORDER BY VALOR DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then 
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRANRIESGOS(POSI,0) = RS(0) 'CODIGO
          TABRANRIESGOS(POSI,1) = RS(1) 'VALOR
          TABRANRIESGOS(POSI,2) = RS(2) 'DESCRIPCION
          TABRANRIESGOS(POSI,3) = RS(3) 'COLOR
 
          TTABRANRIESGO = POSI

        RS.MOVENEXT
     LOOP  



  END IF



     if session("ORPERIODO") = "" then

        Strsql = "Select * From ORPERIODOS  where ESTADO IN ('OK','CL')  " 
        Strsql = Strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
        END IF 

     END IF









    xTrColor = "#FFFFFF"

       wkstylo1 = " style='background-color:#F7CC88;color:#ffffff;' " 

  sImageva    = "<font face='corflaticon' size=2 color='#990000'>&#0110;</font>"
  sImagevm    = "<font face='corflaticon' size=2 color='#f7aa22'>&#0110;</font>"
  sImagevb    = "<font face='corflaticon' size=2 color='#009900'>&#0110;</font>"
  sImagevn    = "<font face='corflaticon' size=2 color='#dddddd'>&#0110;</font>"

  sImageresp    = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:pointer;' >&#0128;</font>"

  sImagePROC    = "<font face='corflaticon' size=2 color='#888888' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:pointer;' >2</font>"

  sImagePROY    = "<font face='corflaticon' size=2 color='#6699cc' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:pointer;' >&#0204;</font>"

  sImageFNL     = "<font face='corflaticon' size=3 color='#009900' style='border-style:OUTSET;border-width:1;background-color:#F7eedd;' >&#097;</font>"

  'Encabezados
  Formulario = ""

    Celda = ""

    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7CC88;'  ","","1%","","","","PLAN","si")

    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7CC88;'  ","","75%","","","","DESCRIPCION","si")
    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7DD88;'  ","","5%","","","","V.RIESGO","si")
    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7DD88;'  ","","5%","","","","COSTO","si")
    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7DD88;'  ","","5%","","","","FECHA","si")
    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7EE88;'  ","","5%","","","","ESTADO","si")
    Formulario = ofv.genrow("","","","","25","","",Celda,"si")


  Strsql = ""



  Strsql = Strsql & "  Select a.*,  "
  Strsql = Strsql & "  f.probabilidad,f.impacto, f.riesgobruto,f.riesgoresidual,f.riesgoresidualNV, "
  Strsql = Strsql & "  x1.porcentaje, x2.impdesde, x2.imphasta, "

  Strsql = Strsql & "   ( ((((x2.impdesde + x2.imphasta)/2) * x1.porcentaje / 100)) / (f.riesgobruto / f.riesgoresidual) ) "
  Strsql = Strsql & "  as ordenador "


'''  Strsql = Strsql & "  ( (((x2.impdesde + x2.imphasta)/2) * x1.porcentaje / 100)  ) as ordenador "

  Strsql = Strsql & "  From ORAGENDA a "
  Strsql = Strsql & "  inner join oractivriesgo f on (a.activriesgo = F.CODRIESGO )  "
  Strsql = Strsql & "  inner join ORNIVPROBOCURRENCIA x1 on (x1.objid = F.probabilidad ) "
  Strsql = Strsql & "  inner join ORNIVIMPACTO x2 on (x2.objid = F.impacto )  "

  Strsql = Strsql & "  WHERE a.ESTADO > 0 AND a.ESTADO < 5 "
  Strsql = Strsql & "  Order By ordenador desc, f.riesgobruto desc, x1.porcentaje desc, a.FECCUMPLI DESC, a.ESTADO "




'''  Strsql = "Select * From ORAGENDA WHERE ESTADO > 0 AND ESTADO < 5 Order By FECCUMPLI DESC, ESTADO "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    WSPVALORT = 0

    IF (RS(16) & "") <> "" THEN     WSPVALORT = RS(16)








  Strsql = "     select d.secuen,e.descri,g.descri, f.probabilidad,f.impacto,"
  Strsql = Strsql & "  f.riesgobruto,f.riesgoresidual,f.riesgoresidualNV,F.CODRIESGO, F.PERIODO,"
  Strsql = Strsql & "   a.* "
  Strsql = Strsql & "  from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & "     inner join objlnk c on (a.cobj = c.codobjeto)"
  Strsql = Strsql & "     inner join objaccion d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced)"
  Strsql = Strsql & "     inner join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)"
  Strsql = Strsql & "     left join oractivriesgo f on (d.codaccion = f.actividad )"
  Strsql = Strsql & "     left join orimpriesgo g on (f.itemriesgo = g.objid)"
  Strsql = Strsql & " where F.CODRIESGO = '" & rs(7) & "' "


''  response.write strsql & "<br>"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then


      WKIPROCD = ""

       FOR WKI=10 TO 17 

           WKIPROC = RS1(WKI) & ""

           IF WKIPROC <> "" THEN
              WKIPROCD = WKIPROCD & RS1(WKI) & " / "
           END IF 


       NEXT

       IF LEN(WKIPROCD) > 4 THEN
          WKIPROCDA = LEFT(WKIPROCD,LEN(WKIPROCD)-3)
       END IF



       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","PROCESO&nbsp;","si")
       sInput = WKIPROCDA
       Celdap = Celdap & ofv.GenCelda("","UtI "  & wkstyloD,"","","","","",sInput,"si")

       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","ACTIVIDAD&nbsp;","si")
       sInput = RS1(0) & "." & RS1(1)
       Celdap = Celdap & ofv.GenCelda("","UtI " & wkstyloD,"","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")

       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","RIESGO&nbsp;","si")
       sInput = RS1(2)
       Celdap = Celdap & ofv.GenCelda("","UtI " & wkstyloD,"","","","","",sInput,"si")

       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","PERIODO&nbsp;","si")

       Speriododsc = RS1(9)

        Strsql = "Select * From ORPERIODOS  where OBJID = 0" & RS1(9)
        set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rsX1.eof then
           Speriododsc              = RSX1(1)

        end if

       sInput = Speriododsc
       Celdap = Celdap & ofv.GenCelda("","UtI " & wkstyloD,"","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")


             wspporc = 100

             Strsql  = "Select VALOR,porcentaje,obs1,DESCRI From ORNIVPROBOCURRENCIA where OBJID = 0" & RS1(3)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then
                wspDATPX      = "0" & RSX1(0)    
                wspDATP       = CINT(wspDATPX)    
                wspporc       = CINT("0" & RSX1(1))
                wspDATprodsc  = RSX1(3) & "&nbsp-&nbsp;" & RSX1(1) & "&nbsp;%&nbsp-&nbsp;" & RSX1(2) 
             END IF

             wspmonto = 0

             Strsql  = "Select VALOR,impdesde,imphasta,obs1,DESCRI From ORNIVIMPACTO where OBJID = 0" & RS1(4)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then
                wspDATIX      = "0" & RSX1(0)    
                wspDATI       = CINT(wspDATIX)    
                wspmonto      = clng((Clng("0" & RSX1(1)) + Clng("0" & RSX1(2))) / 2)
                wspDATimpdsc  = RSX1(4) & "&nbsp-&nbsp;$&nbsp;" & wspmonto & "&nbsp;-&nbsp;" & RSX1(3)
             END IF

             wspmontobruto = CDBL(wspmonto * wspporc / 100)
             wspDATA       = wspDATP * wspDATI   
             wspDATB       = CDBL("0" & RS1(6))  
             wspDATBB      = CDBL("0" & RS1(7))  
             wspDATC       = CDBL(wspDATA / wspDATB)  
             wspDATCN      = CDBL(wspDATA / wspDATBB)  

             Strsql  = "Select descri From ORNIVRIESGOBRUTO where NIVPROB =  0" & RS1(3) & " AND NIVIMPA = 0" & RS1(4) & " "
 '''  response.write strsql & "<br>" & chr(13)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then

                FOR II=1 TO TTABRANRIESGO

                    IF RSX1(0)  =  TABRANRIESGOS(II,0) THEN 
                       wspDATAdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATAdsccl    = TABRANRIESGOS(II,3) 'COLOR
                       EXIT FOR
                    END IF

                NEXT 


             END IF

             wspDATAdsc    = wspDATAdsc & "&nbsp-&nbsp;$&nbsp;" & wspmontobruto

             Strsql  = "Select descri,REQPLAN From ORNIVriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " < VALORDESDE) "
 ''  response.write strsql & "<br>" & chr(13)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then

                FOR II=1 TO TTABRANRIESGO

                    IF RSX1(0)  =  TABRANRIESGOS(II,0) THEN 
 ''  response.write RS1(0) & "<br>" & chr(13)
                       wspDATArrdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATArrdsccl    = TABRANRIESGOS(II,3) 'COLOR
                       EXIT FOR
                    END IF

                NEXT 


             END IF

             Strsql  = "Select descri,REQPLAN From ORNIVriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " < VALORDESDE) "
 ''  response.write strsql & "<br>" & chr(13)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then

                FOR II=1 TO TTABRANRIESGO

                    IF RSX1(0)  =  TABRANRIESGOS(II,0) THEN 
 ''  response.write RS1(0) & "<br>" & chr(13)
                       wspDATArNdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATArNdsccl    = TABRANRIESGOS(II,3) 'COLOR
                       EXIT FOR
                    END IF

                NEXT 


             END IF

             wspvalor      = wspmontobruto / wspDATC
             wspvalorN     = wspmontobruto / wspDATCN

             wspvalor      = ROUND(wspvalor,2)
             wspvalorN     = ROUND(wspvalorN,2)

             wspDATArrdsc    = wspDATArrdsc & "&nbsp-&nbsp;$&nbsp;" & wspvalor
             wspDATArNdsc    = wspDATArNdsc & "&nbsp-&nbsp;$&nbsp;" & wspvalorN







       sInput = Filan

  else

     sInput = "error al buscar la actividad"

  end if


      Celda = ""

    sAccion = ""


    'Descripcion
    sInput = rs(0) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","UT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","RIGHT","","","","",sInput,"si")



    'Descripcion
    sInput = "<B>" & rs(1) & "</B><BR><I>PROCESO:&nbsp;" & WKIPROCDA & "</I>"
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    IF (WSPVALOR & "") > "0" THEN WSPVALOR = Formatnumber(WSPVALOR ,0,0,0,-1)
    IF (WSPVALORT & "") > "0" THEN WSPVALORT = Formatnumber(WSPVALORT ,0,0,0,-1)


    sInput = WSPVALOR & "&nbsp;$&nbsp;"
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","RIGHT","","","","",sInput,"si")


    sInput = WSPVALORT & "&nbsp;$&nbsp;"
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","RIGHT","","","","",sInput,"si")




 

    'FECHA CUMPLIMIENTO


  
    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(2))
    end if

    sInput = dFechan

    sInput = sinput

    Celda = Celda & ofv.gencelda("","VTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' " ,"","","","","",sInput,"si")



    DHOY = CLNG(DATE()) - clng(cdate("01-01-1900"))
    DFEC = clng(rs(2))

    'estado

    DFEC = rs(5) & ""

    if DFEC = "0" then
       sInput = "CANCELADO"
    ELSEIF DFEC = "1" then
       sInput = "PLANIFICADO"
    ELSEIF  DFEC = "2" then
       sInput = "EN CURSO"
    ELSEIF  DFEC = "5" then
       sInput = "FINALIZADO"
    ELSE
       sInput = "REPLANIFICADO"
    END IF


    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Formulario = Formulario & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")

   Celda = ofv.gencelda("","UT COLSPAN=6 ","","","","","","&nbsp;","si")
   Formulario = Formulario & ofv.genrow("","AAA  ","","","","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0" 
       ELSE
          xTrColor = "#FFFFFF" 
       END IF

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")




  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""
  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)


  sHTML = sHTML & GTCCALENDAR("M")





  call ofv.cerrarconn(cn)



  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



Function FechaHAB(wkfecha)

      
      swd = datepart("w",cdate(wkfecha + clng(cdate("01-01-1900"))))

      IF SWD > 6 THEN 
         wkfecha = wkfecha + 2
      END IF

      IF swd = 1 THEN 
         wkfecha = wkfecha + 1
      END IF
    

      FechaHAB  = wkfecha

End Function




%>