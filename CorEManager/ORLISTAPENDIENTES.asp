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
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  SPROCMOD = REQUEST.QUERYSTRING("PROCMOD")

  SPROCESO = REQUEST.QUERYSTRING("PROCESO")
  SFILTRO  = REQUEST.QUERYSTRING("FILTRO")
  SFILSEL  = REQUEST.QUERYSTRING("FILSEL")
  SsubFILTRO  = REQUEST.QUERYSTRING("SUBFILTRO")
  SsubFILTRO2 = REQUEST.QUERYSTRING("SUBFILTRO2")
  SsubFILTRO3 = REQUEST.QUERYSTRING("SUBFILTRO3")
  Speriodo = REQUEST.QUERYSTRING("periodo")


  if Speriodo = "" then Speriodo = session("ORPERIODO")

  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

     if Speriodo = "" then

        Strsql = "Select * From ORPERIODOS  where APLIC = '" & wkaplic & "' and ESTADO IN ('OK','CL')  " 
        Strsql = Strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"

        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
           Speriodo                 = session("ORPERIODO")
           Speriododsc              = session("ORPERIODODESC")

        end if

      else

        Strsql = "Select * From ORPERIODOS  where OBJID = 0" & Speriodo
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           Speriododsc              = RS(1)

        end if

      end if 

  Fila = ""
  xTrColor = "#F9F9F9"


  sImagevMa    = "<font face='arial' size=1 xcolor='#aa0000' style='font-weight:900;background-color:#aa0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MUY&nbsp;ALTO&nbsp;</font>"
  sImageva    = "<font face='arial' size=1 xcolor='#f7aa22' style='font-weight:900;background-color:#f7aa22;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ALTO&nbsp;</font>"
  sImagevm    = "<font face='arial' size=1 xcolor='#f7EEDD' style='font-weight:900;background-color:#f7EEDD;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MODERADO&nbsp;</font>"
  sImagevb    = "<font face='arial' size=1 xcolor='#6699CC' style='font-weight:900;background-color:#6699CC;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ESCASO&nbsp;</font>"
  sImagevNUL    = "<font face='arial' size=1 xcolor='#00AA00' style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NULO&nbsp;</font>"

  sImageSI    = "<font face='arial' size=1  style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;SI&nbsp;</font>"
  sImageNO    = "<font face='arial' size=1  style='font-weight:900;background-color:#AA0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NO&nbsp;</font>"

  sImagevn    = "<font face='arial' size=1 xcolor='#dddddd' style='font-weight:900;background-color:#dddddd;color:#ffffff' >----</font>"

  wkstylof = "AAA style='font-family:arial;font-size:6pt;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' "
  wkstylof2 = "UTI style='border-style:solid;border-width:1;border-color:#aaaaaa;' "


  Strsql = "     select a.cobj,a.cobj, a.cobj, a.cobj,  "
  Strsql = Strsql & " a.cobj, a.cobj, a.cobj,"
  Strsql = Strsql & " a.cobj, a.cobj, a.cobj, a.cobj, a.cobj, A.*, W.ESTADO, W.OBJID AS CODRSKPER "
  Strsql = Strsql & "     from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & "     inner join orrskcircuito w on (a.cobj = w.codobjeto AND W.ESTADO <> 3 AND W.ESTADO <> 4)"



    WKSPSELACC = ""



  Strsql = Strsql & " where " & WKSPSELACC & " w.periodo = 0" & Speriodo & "   "


  Speriododsc = REPLACE(Speriododsc & ""," ","&nbsp;")


 ''' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then 

           Celda = ""


           Celda = Celda & ofv.gencelda("","VTG colspan=7 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY CIRCUITOS PENDIENTES DE CONFIRMACION","si")
           Celda = Celda & ofv.gencelda("","VT  style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"center","","","","","PERIODO:","si")
           Celda = Celda & ofv.gencelda("","VT colspan=3 style='background-color:#f7eedd;color:#888888' " & wkstylo1,"center","","","","",Speriododsc ,"si")

           Fila  = FILA & ofv.genrow("","","","","50","","",Celda,"si")


  ELSE
  posI = 0

  wkstylo1 = " style='border-style:solid;border-width:1;border-color:#aaaaaa;' "


        WKIPROCc = 0
        MAINPROC = 0
        MAINACTI = ""

           Celda = ""
           Celda = Celda & ofv.gencelda("","VTI colspan=4 style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"","","","","","CIRCUITOS PENDIENTES DE CONFIRMACION","si")
           Celda = Celda & ofv.gencelda("","VT  style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"center","","","","","PERIODO:","si")
           Celda = Celda & ofv.gencelda("","VT colspan=6 style='background-color:#f7eedd;color:#888888' " & wkstylo1,"center","","","","",Speriododsc ,"si")

           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")

           Celda = ""

           Celda = Celda & ofv.gencelda("","UTI COLSPAN=7 style='background-color:#8888ff;color:#ffffff' " & wkstylo1,"","","","","","PROCESO","si")
           Celda = Celda & ofv.gencelda("","UTI COLSPAN=4 style='background-color:#8888ff;color:#ffffff' " & wkstylo1,"","","","","","ESTADO","si")
           Fila  = FILA & ofv.genrow("","aaa bgcolor='#f9f9f9' ","","","30","","",Celda,"si")


  Do While (Not rs.EOF) 

        WKIACTIV = ""

        WKIPROCD = ""
        WKIPROCN = 0
        WKIPROCC = RS(27) & ""

        FOR WKI=12 TO 19 

            WKIPROC = RS(WKI) & ""

            IF WKIPROC <> "" THEN
               WKIPROCD = WKIPROCD & RS(WKI) & " / "
            END IF 


        NEXT

        IF LEN(WKIPROCD) > 4 THEN
           WKIPROCD = LEFT(WKIPROCD,LEN(WKIPROCD)-3)
        END IF


        WKIACTIV = WKIPROCC & RS(0) & RS(1) & ""

        IF MAINPROC <> WKIPROCC THEN                     ' CAMBIA PROCESO

         IF MAINPROC <> 0 THEN



         END IF


           Celda = ""
           Celda = Celda & ofv.gencelda("","uti COLSPAN=7 " & wkstylo1,"","","","","",WKIPROCD,"si")
           Celda = Celda & ofv.gencelda("","uti COLSPAN=4 " & wkstylo1,"","","","","",GTESTADODSC(RS("ESTADO")),"si")
           Fila  = FILA & ofv.genrow("","aaa bgcolor='" & xTrColor & "' ","","","30","","",Celda,"si")


           MAINPROC = WKIPROCC

        end if






       if xTrColor = "#FFFFFF" then
          xTrColor = "#F9F9F9" 
       ELSE
          xTrColor = "#FFFFFF" 
       END IF




 

    posi = posi + 1
    rs.Movenext
  Loop

           Celda = ""
           Celda = Celda & ofv.gencelda("","vt","","1%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","40%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","20%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","20%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Fila  = FILA & ofv.genrow("","","","","","","",Celda,"si")


  END IF
  sHTML = sHTML & ofv.GenTabla("","aaa ","","95% ","","0","","0","-1","","",Fila,"si")

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""
  
  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCOBSRSK()

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>