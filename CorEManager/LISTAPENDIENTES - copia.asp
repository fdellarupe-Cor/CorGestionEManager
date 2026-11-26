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

  SPROCESO = REQUEST.QUERYSTRING("PROCESO")
  SFILTRO  = REQUEST.QUERYSTRING("FILTRO")
  SFILSEL  = REQUEST.QUERYSTRING("FILSEL")
  SsubFILTRO  = REQUEST.QUERYSTRING("SUBFILTRO")
  SsubFILTRO2 = REQUEST.QUERYSTRING("SUBFILTRO2")
  SsubFILTRO3 = REQUEST.QUERYSTRING("SUBFILTRO3")

  sMenu = request.querystring("Menu")
  sNombre = request.querystring("Nombre")

  if  SPROCESO<> "" then
  session("SPROCESO")=SPROCESO
  else
  SPROCESO=session("SPROCESO")
  end if

  if SsubFILTRO <> "" then
  session("SsubFILTRO")=SsubFILTRO
  else
  SsubFILTRO=session("SsubFILTRO")
  end if

  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  STRSQL = "SELECT TIPOBJETO  FROM TIPOBJETO WHERE APLIC IN ('EM','GLB')"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     DELIM=""
     STIPOS=""
     DO WHILE NOT RS.EOF

     STIPOS=STIPOS & DELIM & "'" & RS(0) & "'"
     DELIM=","
     RS.MOVENEXT
     LOOP

   STIPOS= " IN(" & STIPOS & ")"


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

''''b.codvinobj = 'VN' AND

      set cnx = ofv.conectar(ofv.strconn5)

      Strsql = ""

      Strsql = Strsql & " SELECT *,cast('12/12/2050' as datetime) - (cast(substring(CKTVIEW.segcktpsc,10,5) as integer)  ) as FECCOMIENZO,cast( ( getdate() - ( cast('12/12/2050' as datetime) - (cast(substring(CKTVIEW.segcktpsc,10,5) as integer)  ) )  ) as integer) as DIFER  FROM CKTVIEW "

	if session("EMPMUN") = "S" then

		Strsql = Strsql & " inner join    (select distinct codempre,codigo,segckt from segcktresp  ) b "
		Strsql = Strsql & " on (b.codigo=cktview.pscresponsable and b.segckt=cktview.segckt) "

	end if

      Strsql = Strsql & " WHERE OBJTIPO " & STIPOS &" AND "

 if cstr(sproceso) <> "0" then

      Strsql = Strsql & "  objtipo ='" & sproceso & "'  AND "

 end if

  if cstr(SsubFILTRO) <> "0" then

       Strsql = Strsql & "  PSCRESPONSABLE ='" & SsubFILTRO & "'  AND "

 end if


      Strsql = Strsql & " FIN = 'NO' AND PSCFIN = 'NO' "

     if session("EMPMUN") = "S" then

		Strsql = Strsql & " and codempre=" & session("UN")

     end if
      Strsql = Strsql & " ORDER BY cktview.SEGCKT DESC, PSCRESPONSABLE, PSCCOMIENZO "

     'response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
  if rs.eof then

           Celda = ""
           Celda = Celda & ofv.gencelda("","VTG colspan=9 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY CIRCUITOS PENDIENTES PARA<BR>LA BUSQUEDA SOLICITADA","si")
           Fila  = FILA & ofv.genrow("","","","","50","","",Celda,"si")


  ELSE
  posI = 0

  wkstylo1 = " style='border-style:solid;border-width:1;border-color:#aaaaaa;' "


        WKIPROCc = 0
        MAINPROC = 0
        MAINACTI = ""

           Celda = ""
           Celda = Celda & ofv.gencelda("","VTI colspan=7 style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"","","","","","LISTA DE CIRCUITOS PENDIENTES DEL WORKFLOW","si")
           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")

           Celda = ""
           Celda = Celda & ofv.gencelda("","uti " & wkstylo1,"","","","","","USUARIO RESPONSABLE","si")
           Celda = Celda & ofv.gencelda("","uti " & wkstylo1,"","","","","","CIRCUITO","si")
           Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","","","","","PASO","si")
           Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","","","","","DOCUMENTO","si")
           Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","","","","","DEMORA / DIAS","si")
           Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","","","","","INICIO","si")
           Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","","","","","REASIGNAR&nbsp;PASO","si")
           Fila  = FILA & ofv.genrow("","aaa bgcolor='#ddddff'  ","","","25","","",Celda,"si")
   if not rs.eof then rs.move(CantRegMoverx)
  Do While (Not rs.EOF) AND POSI < CantRegMostrarx

    sInput2 = rs("FECCOMIENZO") 'cdate(rs(10)&"")

    sinput3 = rs("DIFER") 'clng(date()) - clng(sInput2)

    sinputxx = clng(date())
    sinputzz = clng(sInput2)
    sinputyy = date()




 ''   response.write " Sinput2a ==> " & Sinput2
 ''   response.write " sinput3 ==> "  & sinput3
 ''   response.write " sinputxx ==> " & sinputxx
 ''   response.write " sinputzz ==> " & sinputzz
 ''   response.write " sinputyy ==> " & sinputyy


    smarca = "S"

   ' if ssubfiltro > 0 then

    '   smarca = "N"
'
    '   if ssubfiltro = 99 and sinput3 > 10 then
    '      smarca = "S"
    '   else
    '     if not (sinput3 < ssubfiltro) then
    '        smarca = "S"
    '     end if
'
    '   end if

    'end if


  ' if smarca = "S" then


'''#F7DD88




    Celda = ""

    sInput = RS(14)

sql = "select distinct codempre from segcktresp where segckt='" & rs(5) & "' and codigo='" & rs(6) & "'"
'response.write sql
set rx = ofv.crearconsultaEx(sql,cnX,1,parametros)

codempre=0
if not rx.eof then codempre = rx(0)

    xxxx = getInfoUserEX(RS(14),codempre)
    if session("SGU_nombre") <> "" then    sInput =  session("SGU_nombre")


    Celda = Celda & ofv.gencelda("","UTi  " & wkstylo1 & " BGCOLOR='#F7DD88' ","","","","","",sInput,"si")

    sInput = RS(7)
    Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")

    sInput = RS(16)
    Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")


    sInput = RS(4)

''    response.write " sInput ==> " & sInput

      Strsql = ""

      Strsql = Strsql & " SELECT A.CODOBJETO, (A.DESCRI + ': ' + A.OBJDESCRI) AS NOMBRE "
      Strsql = Strsql & " FROM CSOBJETOSRDO A "
      Strsql = Strsql & " where a.codobjeto = 0" & RS(4)
      Strsql = Strsql & " ORDER BY NOMBRE "
'   response.write strsql & chr(13) & "<br>"
      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      IF NOT rsX1.eof THEN sInput = RSX1(1)
      call ofv.cerrarconsulta(rsX1)

    Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")



    sinput = sinput3

    Celda = Celda & ofv.gencelda("", "UTi " & wkstylo1,"","","","","",sInput,"si")

    sInput =sInput2 'fechaArg(CDate(CLNG(sInput2) + CLng(CDate("01-01-1900"))))' GTFechador(CLNG(sInput2))

''    response.write " Sinput2b ==> " & Sinput2

    Celda = Celda & ofv.gencelda("", "UTi " & wkstylo1,"","","","","",sInput,"si")

     sInput =  "<A HREF='WKFREASIGNADOR.ASP?WKFCOD=" & RS(8) & "&CantRegMoverx=" & CantRegMoverx &"' "
	 sInput =  sInput & "CLASS='NT' TITLE='REASIGNAR PASO A OTRO USUARIO.' ><U>REASIGNAR</U></A>"
	  celda = Celda & ofv.gencelda("", "UTi " & wkstylo1,"","","","","",sInput,"si")



    'Celda = Celda & ofv.gencelda("","UTi  " & wkstylo1 & "  ","","","","","",sInput,"si")

    Fila = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")



       if xTrColor = "#FFFFFF" then
          xTrColor = "#F9F9F9"
       ELSE
          xTrColor = "#FFFFFF"
       END IF


  '  end if



    posi = posi + 1
    rs.Movenext
  Loop

           Celda = ""
           Celda = Celda & ofv.gencelda("","vt","","25%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","25%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","20%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","20%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")


  END IF
  sHTML = sHTML & ofv.GenTabla("","aaa ","","95% ","","0","","0","-1","","",Fila,"si")

  'Botonera
  Agregar = ""
  Buscar  = ""

  sAccion = "SUBCONSULTA.asp?Nombre=" & sNombre & "&Menu=" & sMenu
  Celda = ""'ofv.BotonVolver(sAccion)


  'sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
 sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cnx)

  sHTML = sHTML & "</body></html>"
end if
Function fechaArg(f)
    dia = Day(f)
    mes = Month(f)
    ano = Year(f)
    fechaArg = dia & "/" & mes & "/" & ano
End Function
Response.Write sHTML

%>