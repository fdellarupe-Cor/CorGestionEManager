<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/reportes.inc" -->
<%
Response.buffer=true
SERVER.SCRIPTTIMEOUT=60*8
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

  response.expires=0

   set cn = ofv.conectar(ofv.strconn2)

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)



   xbgcolor=Session("CMFPNIT")
   
     'Encabezados
	 
	set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
	Set oPrtG.Ofv = Ofv
	oPrtG.RCliLogo = session("clilogo")
	oPrtG.RCliNomb = ofv.convertircar("NO SELECCIONADOS","_"," ")
    response.write  oPrtG.RepHeader2(cstr(""), cstr("600"))
	response.write corstyler()


        'response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
        'response.write "<tr>"
        'response.write "<td width=100% >"
        'response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
        'response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
        'response.write "<tr>"
        'response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
        'response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
        'response.write "</tr></table>"

   'Encabezados
   
  'strimestre= request.FORM("TRIMESTRE")



 ' sHTML = ofv.FormHeader(session("FormN"))

  LOTE=REQUEST.QUERYSTRING("LOTE")

  STRSQL = "SELECT a.*,b.*, c.objid as idlote ,c.trimestre,c.ano ,d.trcusrreg as responsable, TRCIMPBRUTO AS MONTO FROM ORRIERRLOTE a"
  STRSQL = STRSQL & " LEFT JOIN ORERRVALIDARI B ON (A.NROERROR = B.CODIGO)"
  STRSQL = STRSQL & " INNER JOIN ORREGINFOBCRA c ON (c.codlote = a.codlote)"
  STRSQL = STRSQL & " INNER JOIN ORtracer d ON (A.NROEVENTO = d.objid)"
  STRSQL = STRSQL & " WHERE A.CODLOTE = '" & LOTE & "' ORDER BY A.NROEVENTO,A.NROERROR DESC"
'  RESPONSE.WRITE STRSQL
  SET RS=OFV.crearconsultaEx(StrSql,cn,1,volver)





   if not rs.eof then
   sInput = "DETALLE DEL LOTE :&nbsp;<b>"& rs("idlote")&"</b> &nbsp;&nbsp;&nbsp;&nbsp;TRIMESTRE :&nbsp;<b>"& rs("TRIMESTRE") & "</b>&nbsp;&nbsp;&nbsp;AÑO :&nbsp;<b>"& rs("ANO")& "</b>"
   else
   sInput ="NO hay EVENTOS SELECCIONADOS"
   end if
   Celda = ofv.gencelda("","CCw " & " BGCOLOR='"&xbgcolor&"'","","90%","","","",sInput,"si")


   Fila = ofv.genrow("","","","","","","",Celda,"si")

   Formulario = Fila
  response.write  ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si") & "<BR>"
  Fila=""



   Formulario = chr(13) & "<TR >"
   Formulario = Formulario & "<TD  class=tt width=5% >ID&nbsp;EVENTO</TD>"
   Formulario = Formulario & "<TD  class=tt width=15% >RESPONSABLE</TD>"
   Formulario = Formulario & "<TD  class=tt width=5% >MONTO</TD>"
   Formulario = Formulario & "<TD  class=tt width=5% >DATO</TD>"
   Formulario = Formulario & "<TD  class=tt width=5% >CODIGO</TD>"
   Formulario = Formulario & "<TD  class=tt width=40% >CAUSA</TD>"
   Formulario = Formulario & "</tr>" & chr(13)

   eventant = 0

   repo = "no"

   filaE   =  ""
   filaN   =  ""

   totevE   =  0
   totevN   =  0

   sumevE   =  0
   sumevN   =  0



  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  FILA=""
  xantevt="0"
  WWWMUESTRAERRF = "S"

     Do While (Not rs.EOF)



     'ID EVENTO
     sInput = RS("NROEVENTO")
     if RS("NROEVENTO") = xantevt then

        cxstyle = "CCs STYLE='border-top-style:none;border-bottom-style:none;' "
        Celda   =  ofv.gencelda("",cxstyle ,"","10%","","","","&nbsp;","si")
     else

         WWWMUESTRAERR = "S"
         WWWMUESTRAERRF = "S"
         IF xantevt <> "0" THEN
            cxstyle = "UT COLSPAN=6 STYLE='BACKGROUND-COLOR:#E4E4E4;FONT-SIZE:6pt;'"
            Celda   =  ofv.gencelda("",cxstyle ,"","","","","","&nbsp;","si")

            if repo = "ENE" then
               filaE = filaE & ofv.genrow("","","","","","","",Celda,"si")
            else
               filaN = filaN & ofv.genrow("","","","","","","",Celda,"si")
            end if 

         END IF

         xantevt = RS("NROEVENTO")
         cxstyle ="ccc STYLE='border-top-style:solid;border-width:1;border-bottom-style:none;text-align:CENTER;'"    ';background-color:" &  Session("CMFPNIT")  & ";'"
         Celda   =  ofv.gencelda("",cxstyle ,"","10%","","","","<B>" & sInput & "</B>","si")




         if RS("NROERROR") > 999 then
            repo = "ENE"
            if RS("NROERROR") = 1001 then
               WWWMUESTRAERR = "N"
            END IF 
         
         else
            repo = "NOR"
         end if 

         if repo = "ENE" then

            totevE   =  totevE + 1
            sumevE   =  sumevE + CDBL(RS("MONTO"))
         else

            totevN   =  totevN + 1
            sumevN   =  sumevN + CDBL(RS("MONTO"))
         end if 





     end if




     sInput = RS("responsable")
     xxx    = getinfouser(sInput & "")

     Celda =  CELDA & ofv.gencelda("","CCx STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:left;' "  ,"","","","","",REPLACE(session("SGU_nombre")&""," " , ""),"si")

     SINPUT = RS("MONTO")
'''''''     Celda  = CELDA & ofv.gencelda("","CCx STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' "  ,"","","","","","&nbsp;$" & SINPUT &"&nbsp;" ,"si")
     Celda  = CELDA & ofv.gencelda("","CCx STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' "  ,"","","","","",SINPUT & "" ,"si")

    'COLUMNA
     sInput = replace(RS("COLUMNA") & ""," ","&nbsp;")
     Celda  = Celda & ofv.gencelda("","CCx STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","","",sInput,"si")

     'NRO ERROR

      sInput = RS("NROERROR")
      IF cint("0" & sInput) >= 1000 then
  	 sInput = "COR-" & SINPUT
      end if
     Celda = Celda & ofv.gencelda("","CCx STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right;' " ,"","","","","",sInput,"si")

     'CAUSA

     wwwcausa =  (RS("CAUSA") & "")

     if wwwcausa = "" then wwwcausa = "ERROR NO TIPIFICADO."

     sInput = REPLACE(wwwcausa,"xxxx","&nbsp;")
     Celda = Celda & ofv.gencelda("","CCx STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","","",sInput,"si")

     if repo = "ENE" then


        IF WWWMUESTRAERRF = "S" THEN

           filaE = filaE & ofv.genrow("","","","","","","",Celda,"si")

        END IF

        WWWMUESTRAERRF = WWWMUESTRAERR


     else
        filaN = filaN & ofv.genrow("","","","","","","",Celda,"si")

     end if 

       posi = posi + 1
       rs.Movenext

    Loop

         IF xantevt <> "0" THEN
            cxstyle = "UT COLSPAN=6 STYLE='BACKGROUND-COLOR:#E4E4E4;FONT-SIZE:6pt;'"
            Celda   =  ofv.gencelda("",cxstyle ,"","","","","","&nbsp;","si")

            if repo = "ENE" then
               filaE = filaE & ofv.genrow("","","","","","","",Celda,"si")
            else
               filaN = filaN & ofv.genrow("","","","","","","",Celda,"si")
            end if 

         END IF

    

    if filaE <> "" then 

       SINPUT = sumevE
       Celda  = ""
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;BACKGROUND-color:#888888;' " ,"","0%","","","","&nbsp;","si") 
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;' " ,"","0%","","","","<B>MONTO</B>","si") 
       Celda  = CELDA & ofv.gencelda("","VTI STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;text-align:right' " ,"","0%","","","",FormatNumber(SINPUT,2) & "","si")
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;' " ,"","0%","","","","<B>EVENTOS</B>","si") 
       Celda  = CELDA & ofv.gencelda("","VTI STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;text-align:right' " ,"","0%","","","",totevE & "&nbsp;","si")
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;BACKGROUND-color:#888888;' " ,"","0%","","","","&nbsp;","si") 
       filaE  = filaE & ofv.genrow("","","","","","","",Celda,"si")

       filaE  = Formulario & filaE
       sHTML  = sHTML & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#000000;' ","","80%  CELLPADDING=0 CELLSPACING=0 bgcolor='#f0f0f0'","","0","","0","0","","",filaE,"si") & "<BR>"
    end if

    if filaN <> "" then 
       SINPUT = sumevN
       Celda  = ""
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;BACKGROUND-color:#888888;' " ,"","0%","","","","&nbsp;","si") 
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;' " ,"","0%","","","","<B>MONTO</B>","si") 
       Celda  = CELDA & ofv.gencelda("","VTI STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;text-align:right' " ,"","0%","","","",FormatNumber(SINPUT,2) & "","si")
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;' " ,"","0%","","","","<B>EVENTOS</B>","si") 
       Celda  = CELDA & ofv.gencelda("","VTI STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;text-align:right' " ,"","0%","","","",totevN & "&nbsp;","si")
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;BACKGROUND-color:#888888;' " ,"","0%","","","","&nbsp;","si") 
       filaN  = filaN & ofv.genrow("","","","","","","",Celda,"si")

       filaN  = Formulario & filaN
       sHTML  = sHTML & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#000000;' ","","80%  CELLPADDING=0 CELLSPACING=0 bgcolor='ddDDff'","","0","","0","0","","",filaN,"si") & "<BR>"
    end if


       SINPUT = sumevE + sumevN
       Celda  = ""
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;color:#FFFFFF;BACKGROUND-color:#000000;' " ,"","0%","","","","&nbsp;TOTAL GENERAL","si") 
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;' " ,"","0%","","","","<B>MONTO</B>","si") 
       Celda  = CELDA & ofv.gencelda("","VTI STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;text-align:right' " ,"","0%","","","",FormatNumber(SINPUT,2) & "","si")
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;' " ,"","0%","","","","<B>EVENTOS</B>","si") 
       Celda  = CELDA & ofv.gencelda("","VTI STYLE='border-style:solid;border-width:1;border-color:#888888;color:#000000;text-align:right' " ,"","0%","","","",totevE + totevN & "&nbsp;","si")
       Celda  = CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;BACKGROUND-color:#888888;' " ,"","0%","","","","&nbsp;","si") 
       filaN  = ofv.genrow("","","","","","","",Celda,"si")
       sHTML  = sHTML & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#000000;' ","","80%  CELLPADDING=0 CELLSPACING=0 bgcolor='BBBBff'","","0","","0","0","","",filaN,"si") & "<BR>"


   ' SELECT

    STRSQL=" SELECT MAX(A.TOTAL),MAX(A.totalMontoTri) , SUM(MONTO),MAX(EVENTORAIZ), MAX(TOTALINFORMAR),MAX(TOTALNOINFORMAR) ,max(totrep) AS totrep,max(totenca) AS totenca,max(totdifrep) AS totdifrep,max(a.trimestre),max(a.ano),MAX(MONTOTOTALNOINF) AS MONTOTOTALNOINF ,max(a.objid) as idlote FROM ORREGINFOBCRA A "
    STRSQL =STRSQL & "INNER JOIN ORREGINFOBCRADET B  ON (A.CODLOTE=B.CODLOTE) "
    STRSQL = STRSQL & " WHERE B.CODLOTE='" &  LOTE &"'"
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    Celda =   ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","0%","","","","<B>TOTAL DE REGISTROS</B>","si") & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right' " ,"","0%","","","",rs2(5),"si")

    if isnumeric(replace(RS2("MONTOTOTALNOINF") & "",",",".")) then

       Celda =   CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","0%","","","","<B>MONTO</B>","si") & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right' " ,"","0%","","","",FormatNumber(RS2("MONTOTOTALNOINF"),2),"si")
    else
       Celda =   CELDA & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","0%","","","","<B>MONTO</B>","si") & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;text-align:right' " ,"","0%","","","","0","si")
    end if

    fila =  ofv.genrow("","","","","","","",Celda,"si")


 '  sHTML = sHTML  & "<BR>" & ofv.GenTabla("","","","80% ","","0","","0","0","","",FILA,"si")
   call ofv.cerrarconsulta(rs2)

   sAccion = "ORGENRUNOR.asp?CantRegMover=" & CantRegMoverX & ""
   Celda = ofv.BotonVolver(sAccion)



  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


end if
  Response.write shtml



%>