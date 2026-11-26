<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/reportes.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

  session("ModMain") = "CL"

  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"

 set cn = ofv.conectar(ofv.strconn2)

 OPC=REQUEST.QUERYSTRING("OPC")

set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")


Set oPrtG.Ofv = Ofv



 sHTML = ofv.FormHeader(session("FormN")) & chr(13)

select case OPC

CASE "LST"
   sHTML=""

   sHTML = corstyler()

 '  response.write ofv.Repbody(request.querystring("mnn"))
   xsHTML=   "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
   xsHTML=  xsHTML & "<tr>"
   xsHTML=  xsHTML & "<td width=100% >"
 '  xsHTML=  xsHTML &  ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
   xsHTML=  xsHTML & "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
   xsHTML=  xsHTML & "<tr>"
   xsHTML=  xsHTML & "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
   xsHTML=  xsHTML & "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
   xsHTML=  xsHTML & "</tr></table>"

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = "MAPA DE PROCESOS"
   xsHTML=  oPrtG.RepHeader2(cstr(""), cstr("600"))


  ORPERIODO    = request.querystring("PERIODO") & ""

  TTMAP = 0
  ITMAP = 0
  CTMAP = 0
  LTMAP = 0

  DIM TABPROCESOSMAP (500,10)


         Strsql2 = "Select OBJID,PERIODO,RGAMBAUTOEV From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & ORPERIODO
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODODSC  = RSX1(1)
            WKORPERIODOMAP  = RSX1(2)
         ELSE
            WKORPERIODODSC  = "NO ENCONTRADO"
            WKORPERIODOMAP  = 0
         END IF



''''' MASTER LOOP DE PROCESOS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      scodmaster = WKORPERIODOMAP
      stipmaster = "PROCESO"         

      Strsql = "Select OBJETOS.* From relaobj, OBJETOS  where CODOJASOC = CODOBJETO "
      Strsql = Strsql & " AND codobjmaster = '" & scodmaster & "' and tipobjmaster = '" & stipmaster & "'  "
      Strsql = Strsql & " ORDER BY OBJETOS.INDOBJ "
    '  RESPONSE.WRITE STRSQL

      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsX1.eof then
         TTMAP = 1
         ITMAP = 1  '   
         ITMAP2 = 1  '        
         DO UNTIL RSX1.EOF

            WKFPROCESO = RSX1(0) 


            Celda =  ofv.gencelda("","UTI style='background-color:#8888ff;color:#ffffff' COLSPAN=4","","1%","","","","PROCESO :&nbsp;" & RSX1(0) & "-" & RSX1(3),"si")
            xFila  = xFILA & ofv.genrow("","aaa bgcolor='#f9f9f9' ","","","","","",Celda,"si")

            Celda =  ofv.gencelda(""," COLSPAN=4","","","","","","&nbsp;","si")
            xFila  = xFILA & ofv.genrow(""," ","","","","","",Celda,"si")

            Celda = ofv.gencelda("","uti  " ,"","","","","","<b>Nº</b>","si")
            Celda = Celda & ofv.gencelda("","UTI " ,"","20%","","","","<b>SUBPROCESO<b/>","si")
            Celda = Celda & ofv.gencelda("","UTI " ,"","","","","","<b>Nº</b>","si")
            Celda = Celda & ofv.gencelda("","UTI " ,"","","","","","<b>ACTIVIDADES</b>","si")

       '     xFila  = xFILA & ofv.genrow("","aaa bgcolor='#F7EEDD' ","","","","","",Celda,"si")

            TABPROCESOSMAP (ITMAP,1) = WKFPROCESO


            strsql =  "select   B.CODOBJETO,B.DESCRI,B.CODVINOBJ, A.*    "
            strsql =  strsql &"     from csrelpathdescri a  "
            strsql =  strsql &"    inner join objetos b on (a.cobj = b.codobjeto)  "
            Strsql =  Strsql & "   inner join objprty n on (b.codobjeto=n.codobjeto and estado=3)"
            strsql =  strsql &"          where      ( "
            strsql =  strsql &"          A.MC1 =" & WKFPROCESO & " OR  A.MC2 = " & WKFPROCESO & " OR  A.MC3 = " & WKFPROCESO
            strsql =  strsql &"          OR  A.MC4 = " & WKFPROCESO & " OR  A.MC5 = " & WKFPROCESO & " OR  A.MC6 = " & WKFPROCESO & " OR  "
            strsql =  strsql &"          A.MC7 = " & WKFPROCESO & " OR  A.COBJ =" & WKFPROCESO & "   ) "
            strsql =  strsql &"          AND (  B.APLIC = 'RG' ) "
            strsql =  strsql &"           ORDER BY  A.TSEC"

            set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if NOT rsX2.eof then
   
               DO UNTIL RSX2.EOF

                  WKFTSEC = "" & RSX2("TSEC")

                  WKFTSECN = LEN(WKFTSEC)

                  WKFTSECN = CINT(WKFTSECN / 6) - 1

                  WKFTSECD = "" & RSX2(1)

                  FOR III=2 TO  WKFTSECN

                      WKFTSECD = "--" & WKFTSECD

                  NEXT                 

                  WKFTSECTAB = SPLIT(WKFTSEC,".")  

                  WKFTSECOBJ = ""
                  FOR IIII=0 TO UBOUND(WKFTSECTAB)
                      WKFTSECOBJ = WKFTSECOBJ & (CLNG(WKFTSECTAB(IIII)) - 900000) & "/"
                  
                  NEXT

                  Celda = ofv.gencelda("","uti  " ,"","5%","","","",RSX2(0),"si")
                  Celda = Celda & ofv.gencelda("","UTI " ,"","70%","","","",UBOUND(WKFTSECTAB) & "-" & WKFTSECOBJ,"si")
                  Celda = Celda & ofv.gencelda("","UTI " ,"","15%","","","",RSX2("TSEC"),"si")
                  Celda = Celda & ofv.gencelda("","UTI " ,"","5%","","","",RSX2(2),"si")

                  xFila  = xFILA & ofv.genrow("","aaa bgcolor='#EEEEEE' ","","","","","",Celda,"si")

                  Iraiz = RSX2(0)

                  strsql =  "select   D.DESCRIACC as actividad , D.CODOBJETO AS CODOBJACC, A.*  ,d.secuen  "
                  strsql =  strsql &" from csrelpathdescri a  "
                  strsql =  strsql &"   inner join objetos b   on (a.cobj = b.codobjeto)  "
                  strsql =  strsql &"   inner join objproced c on ( c.codobjeto = B.codobjeto )  "
                  strsql =  strsql &"   inner join objaccion d on ( C.codproced = d.codproced) "
                  strsql =  strsql &" Where (   A.COBJ =" & Iraiz & " )"
                  strsql =  strsql &"   AND (  B.APLIC = 'RG' )"
                  strsql =  strsql &"   ORDER BY  A.TSEC,D.SECUEN"


                  strsql =  "select   A.COBJ "
                  strsql =  strsql &" from csrelpathdescri a  "
                  strsql =  strsql &"   inner join objetos b   on (a.cobj = b.codobjeto)  "
                  strsql =  strsql &"   inner join objproced c on ( c.codobjeto = B.codobjeto )  "
                  strsql =  strsql &" Where (   A.COBJ =" & Iraiz & " )"
                  strsql =  strsql &"   AND (  B.APLIC = 'RG' )"



                  set rsX3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                  if NOT rsX3.eof then
   
                   '  DO UNTIL RSX3.EOF

                     '   WKFTSECD = "..................." & RSX3(0)

                      '  Celda = ofv.gencelda("","uti  " ,"","5%","","","",RSX3("secuen"),"si")
                      '  Celda = Celda & ofv.gencelda("","UTI " ,"","70%","","","",WKFTSECD,"si")
                      '  Celda = Celda & ofv.gencelda("","UTI " ,"","15%","","","",RSX3(1),"si")
                      '  Celda = Celda & ofv.gencelda("","UTI " ,"","5%","","",""," ","si")


                        Celda = ofv.gencelda("","uti  " ,"","5%","","","",RSX2(0),"si")
                        Celda = Celda & ofv.gencelda("","UTI COLSPAN=3" ,"","70%","","","","NO COPIA","si")
                        xFila  = xFILA & ofv.genrow("","aaa bgcolor='#FF9999' ","","","","","",Celda,"si")


                    '    RSX3.MOVENEXT
                    '  LOOP  

                  ELSE


                        Celda = ofv.gencelda("","uti  " ,"","5%","","","",RSX2(0),"si")
                        Celda = Celda & ofv.gencelda("","UTI COLSPAN=3" ,"","70%","","","","SI COPIA","si")
                        xFila  = xFILA & ofv.genrow("","aaa bgcolor='#99FF99' ","","","","","",Celda,"si")


                  END IF


                  RSX2.MOVENEXT
                LOOP  

             END IF



            RSX1.MOVENEXT
         LOOP  

      END IF

   sHTML=shtml & xshtml &  ofv.GenTabla("","AAA  " & wkstylo1 ,"","100% ","","0","","0","-1","","",xfila,"si") 
   call ofv.cerrarconsulta(RSX2)



case else


    XTPPER      = "" & request.form("PERIODO")



'--------------------
      sAccion = ""


      Celda  = ofv.GenCelda("","tt","","","","","","PERIODO","si")

      Strsql = "Select OBJID,PERIODO,RGAMBAUTOEV, (99999 - FINAL) AS XFFINAL,INICIO, ESTADO From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
      Strsql = Strsql & " ESTADO IN ('OK','CL')  "
      Strsql = Strsql & " ORDER BY XFFINAL DESC,INICIO, ESTADO DESC"

      IF XTPPER = "" THEN

         set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if NOT rsX1.eof then
            XTPPER          = RSX1(0)
            WKORPERIODOMAP  = RSX1(2)
         ELSE
            WKORPERIODOMAP  = 0
         END IF

      ELSE 


         Strsql2 = "Select OBJID,PERIODO,RGAMBAUTOEV From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & XTPPER
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODOMAP  = RSX1(2)
         ELSE
            WKORPERIODOMAP  = 0
         END IF


      END IF


      sInput = ofv.GenerarCombo("Subconsulta0","PERIODO",XTPPER,Strsql,cn,0,1,"")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")

      Fila   = ofv.GenRow("","","","","30","","",Celda,"si")


      sAccion    = "RGGENmapaproceso.asp?Opc=SET"
      Formulario = Formulario & ofv.GenForm("Subconsulta0","",sAccion,"","post","",Fila,"si")


      Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
      sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
      sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
      Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
      formx    = ofv.GenRow("","","","","40","","",Celda,"si")
      sAccion    = "RGGENmapaproceso.asp?Opc=LST&PERIODO=" & XTPPER
      Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"_NEW","post","",formx,"si")


'-----------------------------



   sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;' ","CENTER","60% ","","0","","5","0","","",Formulario,"si")



  sHTML = sHTML & "</center>"



  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
end select
END IF

response.write sHTML

call ofv.cerrarconn(cn)

%>