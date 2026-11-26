<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


                      '      WSPDIVN = "DENC" & xtmsgENCCOD
                      '      WSPFRMN = "FENC" & xtmsgENCCOD & "IT" & rs2(0)
                      '      WSPFRMA = "EMENCUESTASACT.ASP?TP=ES&OP=ACT&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)


 SENCTP   = REQUEST.QUERYSTRING("TP")   ' TIPO DE ENCUESTA -    ES= ENC.SIMPLE   EA= ENC.AVANZADA
 SENCOP   = REQUEST.QUERYSTRING("OP")   ' OPERACION        -    ACT= ACTUAL.VOTO   RES= MUESTRA RESULTADOS SOLAMENTE
 SENCID   = REQUEST.QUERYSTRING("ENC")  ' ID DE LA ENCUESTA
 SENCBLK  = REQUEST.QUERYSTRING("BLK")  ' ID DEL BLOQUE
 SENCITM  = REQUEST.QUERYSTRING("ITM")  ' ID DEL ITEM


 WSPDIVN = "DENC" & SENCID              ' ID DEL DIV DE VENTANA PARA MOSTRAR


 DIM TABENCUESTA (50,50,10)

 TTABENC = 0


 sHTML = sHTML & "<HTML>"
 sHTML = sHTML & "<BODY onload=""window.top.LDVENTANA('" & WSPDIVN & "');"" >"

 DHOY = clng(date()) - clng(cdate("01-01-1900"))

  set cn = ofv.conectar(ofv.strconn2)


 IF SENCOP = "ACT" THEN


        strsql1 = "Insert into EMENCUESTASVTN "
        strsql1 = strsql1 & " (ENCcodigo,BLKcodigo,VTNusugrp,VTNfecVOTO,VTNhora, VTNVALORITM,VTNOBSERVACION"
        strsql1 = strsql1 & " ) Values ("
        strsql1 = strsql1 &  "'" & SENCID & "','" & SENCBLK & "','" & SESSION("USUARIO") & "','" & DHOY & "',"
        strsql1 = strsql1 &  "'" & TIME() & "','" & SENCITM & "','NO'"
        strsql1 = strsql1 & ")"
  '''response.write strsql1 & "<br>"
        call ofv.crearconsultaEx(StrSql1,cn,1,volver)

        strsql = "update EMENCUESTASBLK set BLKTOTAL = (BLKTOTAL + 1) where ENCcodigo = 0" & SENCID
        Strsql = Strsql & " and objid = 0" & SENCBLK
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

        strsql = "update EMENCUESTASITM set ITMTOTAL = (ITMTOTAL + 1) where OBJID = 0" & SENCITM
        Strsql = Strsql & " AND ENCcodigo = 0" & SENCID
        Strsql = Strsql & " and BLKcodigo = 0" & SENCBLK
'''  response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)



 END IF


 
  Strsql = "Select * From emencuestas  where objid = 0" & SENCID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     TABENCUESTA (0,0,0) = SENCID
     TABENCUESTA (0,0,1) = rs(17) & ""


'''''''''''' BLOQUES DE LA ENCUESTA ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     Strsql = "Select * From EMENCUESTASBLK WHERE ENCcodigo = 0" & SENCID
     Strsql = Strsql & " Order By BLKSECUEN,objid "

     set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     posI = 1

     Do While (Not rs1.EOF)

        sblkcodigo             = rs1(0)

        TABENCUESTA (posI,0,0) = rs1(0)
        TABENCUESTA (posI,0,1) = rs1(3) & ""
        TABENCUESTA (posI,0,2) = rs1(4) & ""
        TABENCUESTA (posI,0,3) = 0

        TABENCUESTA (posI,0,4) = rs1(5)



''''''''''''''''''' items del bloque de encuesta ''''''''''''''''''''''''''''''''''''''''''''''''

        Strsql2 = "Select * From EMENCUESTASITM WHERE ENCcodigo = 0" & SENCID
        Strsql2 = Strsql2 & " and BLKcodigo = 0" & sblkcodigo
        Strsql2 = Strsql2 & " Order By ITMSECUEN,objid "
        set rs2 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)

        posi2 = 1
        Do While (Not rs2.EOF)

           TABENCUESTA (posI,posi2,0) = rs2(0)
           TABENCUESTA (posI,posi2,1) = rs2(4) & ""
           TABENCUESTA (posI,posi2,2) = rs2(5) & ""
           TABENCUESTA (posI,posi2,3) = 0
           TABENCUESTA (posI,posi2,4) = rs2(6)

           TABENCUESTA (posI,0,3)     = posi2

           posi2 = posi2 + 1
           rs2.Movenext
        Loop

        TTABENC = posi

        posi = posi + 1
        rs1.Movenext
     Loop

  end if



''''''''''''''''''''''''''''''''''' MUESTRA RESULTADOS ''''''''''''''''''''''''''''''''''''''''''''''

    Fila = ""


    wkproccod = ""

    sImage = "<font face='corflaticon' size=2 color='#f0f0f0' style='background-color:#bb0000;border-style:OUTSET;border-width:2;' >&#0120;</font>"

    wkproccod = "<a class=nt  href='javascript:void(0);' ><b>"
    wkproccod = wkproccod & sImage
    wkproccod = wkproccod & "</b></a>"


    CELDAD = ""
    CeldaD = CeldaD & ofv.GenCelda("","utI  ","","95%","","","","ENCUESTAS","si")
    CeldaD = CeldaD & ofv.GenCelda("","uti title='Cerrar.' onclick="" document.getElementById('" & WSPDIVN & "').style.display='none'; "" ","center","5%","","","",wkproccod,"si")
    FilaD     = ofv.genrow("","","","","","","",CeldaD,"si")
    sFORMhed = ofv.GenTabla("","TWIND","","100%","","0","","0","0","","",FilaD,"si")

    Celda = ofv.GenCelda("","ut colspan=7 ","","","","","",sFORMhed,"si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ""
    Celda = Celda & ofv.GenCelda("","VtI style='FONT-SIZE:8pt;background-color:#F0F0F0;color:#AA0000;' colspan=7 ","","","","","",TABENCUESTA (0,0,1),"si")
    Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")

    CELDA = ""
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

 IF SENCTP = "ES" THEN

    posI = 1

    CELDA = ""
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","10%","","","","&nbsp;","si")
  '  Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    WSPTOTAL = TABENCUESTA (posI,0,4)
    WSPitems = TABENCUESTA (posI,0,3)


''''''''''''''''''' items del bloque de encuesta ''''''''''''''''''''''''''''''''''''''''''''''''
    Filax = ""

    Celda = ofv.GenCelda("","Uti style='font-size:8pt;background-color:#F0F0F0;color:#000000;' COLSPAN=12 ","","","","","",capitalize(TABENCUESTA (posI,0,1) & ""),"si")

    Filax = filax & ofv.genrow("","","","","20","","",Celda,"si")


    CELDA = ""
    Celda = Celda & ofv.GenCelda("","Ut ","","30%","","","","&nbsp;","si")

    for ii=1 to 10
        Celda = Celda & ofv.GenCelda("","Ut","CENTER","7%","","","","&nbsp;","si")
    next

    Celda = Celda & ofv.GenCelda("","Ut ","CENTER","5%","","","","&nbsp;","si")
    Filax = filax & ofv.genrow("","","","","","","",Celda,"si")

       WSPCOLOR = "#ff0000"

       for posi2=1 to  WSPitems

         WSPTOTALitm = TABENCUESTA (posI,posi2,4)

         if WSPTOTALitm = 0 then

            WSPTOTALprc = 0

            WSPTOTALdib = 0  

         else 

            WSPTOTALprc = round((WSPTOTALitm / WSPTOTAL * 100),1)

            WSPTOTALdib = cint(WSPTOTALitm / WSPTOTAL * 10 )      

         end if     

         CELDA = ""

         Celda = Celda & ofv.GenCelda("","vtI ","","","","","",TABENCUESTA (posI,posi2,1) & "&nbsp;(" & WSPTOTALprc & "%)","si")

         for ii=1 to WSPTOTALdib
             Celda = Celda & ofv.GenCelda("","Ut bgcolor='" & WSPCOLOR & "' ","CENTER","","","","","&nbsp;","si")
         next

         for ii=1 to (10 - WSPTOTALdib)
             Celda = Celda & ofv.GenCelda("","Ut bgcolor='#DDDDDD' ","CENTER","","","","","&nbsp;","si")
         next

         Celda = Celda & ofv.GenCelda("","Ut ","CENTER","","","","","&nbsp;","si")

         Filax = filax & ofv.genrow("","","","","20","","",Celda,"si")

         Celda = ofv.GenCelda("","AAA STYLE='font-family:arial;font-size:5pt;' COLSPAN=12 ","CENTER","","","","","&nbsp;","si")

         Filax = filax & ofv.genrow("","","","","5","","",Celda,"si")


         WSPCOLOR = "#0000DD"

      next 

      Celda = ofv.GenCelda("","Ut style='font-size:8pt;color:#888888;' COLSPAN=12 ","RIGHT","","","","",capitalize("TOTAL VOTOS: " & WSPTOTAL & "") & "&nbsp;&nbsp;","si")

      Filax = filax & ofv.genrow("","","","","20","","",Celda,"si")



      wsptabla = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#EEEEEE;' ","","80%","","0","","-2","-1","","",FILAx,"si")


      Celda = ""
      Celda = Celda & ofv.GenCelda("","Ut  colspan=7 ","CENTER","","","","",wsptabla,"si")
      Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")

 ELSE



 END IF

 shtml = shtml & ofv.GenTabla("","AAA  ","","100%","","0","","0","0","","",FILA,"si")


  sHTML = sHTML & "</BODY>"

  sHTML = sHTML & "</HTML>"
end if

 response.write shtml



%>