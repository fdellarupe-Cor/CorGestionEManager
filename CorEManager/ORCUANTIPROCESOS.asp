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
oPrtG.RCliNomb = "CUANTIFICACION DE PROCESOS"
   xsHTML=  oPrtG.RepHeader2(cstr(""), cstr("600"))


  'proceso=request.querystring("proceso")
  'subproceso=request.querystring("suBproceso")

ORPERIODO=REQUEST.FORM("PERIODO")

ORDESDE=REQUEST.FORM("ORDESDE")


ORHASTA=REQUEST.FORM("ORHASTA")

IF ORDESDE="TODOS" THEN ORDESDE=0

IF ORHASTA="TODOS" THEN ORHASTA=9999

IF NOT IsNumeric(ORDESDE) THEN ORDESDE=0

IF NOT IsNumeric(ORHASTA) THEN ORHASTA=9999


RANGO = " AND PRIORIDAD BETWEEN "& ORDESDE &" AND " & ORHASTA



         Strsql2 = "Select OBJID,PERIODO,RGAMBAUTOEV From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & ORPERIODO
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODOMAP  = RSX1(2)
         ELSE
            WKORPERIODOMAP  = 0
         END IF



strsql ="Select OBJETOS.* From relaobj, OBJETOS  where CODOJASOC = CODOBJETO  AND codobjmaster = '0" & WKORPERIODOMAP & "' and tipobjmaster = 'PROCESO' " & RANGO  &   " ORDER BY prioridad,descri"
'response.write strsql
set rsProc = ofv.crearconsultaEx(StrSql,cn,1,parametros)

Iporcesos=0
Iprioridad=""

  Celda = ofv.gencelda("","uti  ALIGN=CENTER" ,"","","","","","<b>PRIORIDAD</b>","si")

  Celda = Celda & ofv.gencelda("","UTI ALIGN=CENTER" ,"","60%","","","","<B>PROCESOS</B>","si")
  Celda = Celda & ofv.gencelda("","UTI ALIGN=CENTER" ,"","","","","","<b>SUBPROCESOS</b>","si")
  Celda = Celda & ofv.gencelda("","UTI ALIGN=CENTER" ,"","","","","","<b>ACTIVIDADES</b>","si")

  xFila  = xFILA & ofv.genrow("","aaa bgcolor='#F7EEDD' ","","","","","",Celda,"si")

  Celda=""
  Iprioridad=""
  Itotal=0
  NOhAY=0

do while not  rsProc.eof
    NOhAY=1
    ReDim Preserve ArrayP(3, Iporcesos)


    ArrayP(0,Iporcesos)=rsProc("descri")
    ArrayP(1,Iporcesos)=0
    ArrayP(2,Iporcesos)=0
    ArrayP(3,Iporcesos)=rsProc("prioridad")




    Iraiz=rsProc(0)


    strsql =  "select   A.EXPR7,COUNT(A.EXPR7) AS cantAct   from csrelpathdescri a  "
    strsql =  strsql &" inner join objetos b on (a.cobj = b.codobjeto)  "
    Strsql = Strsql & "  inner join objprty n on (b.codobjeto=n.codobjeto and estado=3)"
    strsql =  strsql &"      inner join objlnk x on (a.cobj = x.codobjeto and (x.ltipo = 'LRGVN' or x.ltipo = 'LRGCR'))    "
    strsql =  strsql &"        inner join objproced c "
    strsql =  strsql &"   on ( c.codobjeto = x.codobjetol and c.codproced = x.iden4 )  "
    strsql =  strsql &"      inner join objaccion d on ( C.codproced = d.codproced) "
    strsql =  strsql &"          LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt and x.ltipo = 'LRGVN' "
    strsql =  strsql &"          )"
    strsql =  strsql &"          where      (   A.MC6 = "& Iraiz & " OR  A.MC7 = "& Iraiz & " OR  A.COBJ ="& Iraiz & "   ) AND (  B.APLIC = 'RG' )  GROUP BY A.EXPR7  "
    'response.write strsql & "<br>"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do while not rs.eof
   ArrayP(1,Iporcesos)=ArrayP(1,Iporcesos) + 1
   ArrayP(2,Iporcesos)=ArrayP(2,Iporcesos) + rs("cantAct")
   rs.movenext
  loop
  call ofv.cerrarconsulta(rs)
  if Iprioridad <> rsProc("prioridad") then

   ReDim Preserve ArrayT(2,Itotal)
   ArrayT(0,Itotal)=0
   ArrayT(1,Itotal)=0
   ArrayT(2,Itotal)=1
   ArrayT(0,Itotal)=ArrayT(0,Itotal) + ArrayP(1,Iporcesos) ' totales subprocesos del proceso
   ArrayT(1,Itotal)=ArrayT(1,Itotal) + ArrayP(2,Iporcesos) ' totales actividades del procesos
   Iprioridad = rsProc("prioridad")

   Itotal=Itotal+1
   else
    ArrayT(0,Itotal-1)=ArrayT(0,Itotal-1) + ArrayP(1,Iporcesos) ' totales subprocesos del proceso
    ArrayT(1,Itotal-1)=ArrayT(1,Itotal-1) + ArrayP(2,Iporcesos) ' totales actividades del procesos
    ArrayT(2,Itotal-1)=ArrayT(2,Itotal-1) + 1 ' totales actividades del procesos
  end if
  Iporcesos=Iporcesos + 1
  rsProc.movenext

loop
  call ofv.cerrarconsulta(rsProc)


  IF NOhAY=0 THEN
   Celda =  ofv.gencelda("","uti  " ,"","","","","","No exite informacion para su busqueda","si")
    Fila  = FILA & ofv.genrow("","AAA  " ,"","","","","",Celda,"si")
   sHTML=sHTML & ofv.GenTabla("","AAA " ,"","100% ","","0","","0","-1","","",Fila,"si")
  ELSE
  fila=""
  xTrColor = "#EEEEEE"
  Actprioridad=ArrayP(3,0)
  Iprioridad=0
  procesostotales= ArrayT(2,Iprioridad)
  subprocesostotales= ArrayT(0,Iprioridad)
  actividadestotales= ArrayT(1,Iprioridad)
  ilinea=0


  for i= 0 to Iporcesos-1


  if Actprioridad <> ArrayP(3,I) then

   Celdax =  ofv.gencelda("","uti  colspan=2 " ,"","","","","","<font size=2><b> Cantidad :"&ArrayT(2,Iprioridad) & "</b></font>","si")
   Celdax =  celdax & ofv.gencelda("","uti  ALIGN=RIGHT" ,"","","","","","<b><font size=2>" & ArrayT(0,Iprioridad) & "</font></b>","si")
   Celdax =  celdax & ofv.gencelda("","uti  ALIGN=RIGHT" ,"","","","","","<b><font size=2>" & ArrayT(1,Iprioridad)& "</font></b>","si")
   Fila  = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celdax,"si")
   Celdax =  ofv.gencelda("","uti  colspan=4 " ,"","","","","","&nbsp;" ,"si")
   Fila  = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celdax,"si")
   Celdax=  ofv.gencelda("","uti  " ,"","","","","",ArrayP(3,I),"si")
   Celdax=  celdax & ofv.gencelda("","uti  ALIGN=LEFT" ,"","","","","",ArrayP(0,I),"si")
   Celdax =  celdax & ofv.gencelda("","uti  ALIGN=RIGHT" ,"","","","","",ArrayP(1,I),"si")
   Celdax =  celdax & ofv.gencelda("","uti  ALIGN=RIGHT" ,"","","","","",ArrayP(2,I),"si")

   Fila  = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celdax,"si")
   Actprioridad=ArrayP(3,I)
   Iprioridad= Iprioridad+1
   procesostotales=procesostotales + ArrayT(2,Iprioridad)
   subprocesostotales=subprocesostotales + ArrayT(0,Iprioridad)
   actividadestotales=actividadestotales + ArrayT(1,Iprioridad)
   else

     Celda =  ofv.gencelda("","uti  " ,"","","","","",ArrayP(3,I),"si")
     Celda =  celda & ofv.gencelda("","uti  ALIGN=LEFT" ,"","","","","",ArrayP(0,I),"si")
     Celda =  celda & ofv.gencelda("","uti  ALIGN=RIGHT" ,"","","","","",ArrayP(1,I),"si")
     Celda =  celda & ofv.gencelda("","uti  ALIGN=RIGHT" ,"","","","","",ArrayP(2,I),"si")
     Fila  = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
  end if


  if ilinea> 25 then
   sHTML=shtml & xshtml &  ofv.GenTabla("","AAA  " & wkstylo1 ,"","600 ","","0","","0","-1","","",xfila & Fila,"si") & "<DIV style='page-break-before:always'></DIV> "
    fila=""
    ilinea=0
  else
    ilinea=ilinea +1
  end if

  next

     Celda =  ofv.gencelda("","uti  colspan=2 " ,"","","","","","<font size=2><b> Cantidad :"&ArrayT(2,Iprioridad) & "</b></font>","si")
     Celda =  celda & ofv.gencelda("","uti  ALIGN=RIGHT" ,"","","","","","<b><font size=2>" & ArrayT(0,Iprioridad) & "</font></b>","si")
     Celda =  celda & ofv.gencelda("","uti  ALIGN=RIGHT" ,"","","","","","<b><font size=2>" & ArrayT(1,Iprioridad)& "</font></b>","si")
     Fila  = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  sHTML=shtml & xshtml &  ofv.GenTabla("","AAA  " & wkstylo1 ,"","600 ","","0","","0","-1","","",xfila & Fila,"si") & "<DIV style='page-break-before:always'></DIV> "

  fila=""
  celda=ofv.gencelda("","uti  colspan=2 " ,"","","","","","<font size=2><b>Total de procesos :"& procesostotales & "</b></font>","si")
  Celda =  celda & ofv.gencelda("","uti  ALIGN=RIGHT" ,"","","","","","<b><font size=2>Total de subprocesos :" & subprocesostotales & "</font></b>","si")
  Celda =  celda & ofv.gencelda("","uti  ALIGN=RIGHT" ,"","","","","","<b><font size=2>Total de actividades :" & actividadestotales & "</font></b>","si")
  Fila  = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  sHTML=shtml & "<br>" & ofv.GenTabla("","AAA  " & wkstylo1 ,"","600 ","","0","","0","-1","","", Fila,"si") & "<DIV style='page-break-before:always'></DIV> "




  END IF

  'sHTML=sHTML & ofv.GenTabla("","AAA  " & wkstylo1 ,"","100% ","","0","","0","-1","","",Fila,"si")
case else


    scodmaster = "000"
    stipmaster = "PROCESO"

    XTPROC = clng("0" & request.form("PROC"))
    XSUBPROCESO = clng("0" & request.form("SUBPROCESO"))

  '  response.write request.form


'--------------------
      sAccion = "ORCUANTIPROCESOS.ASP?OPC=LST"

      FILA=""

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


      END IF


      sInput = ofv.GenerarCombo(sForm,"PERIODO",XTPPER,Strsql,cn,0,1,"no")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")

      Fila   = ofv.GenRow("","","","","30","","",Celda,"si")







      Celda  = ofv.GenCelda("","tt","","","","","","PRIORIDAD DESDE","si")



      SINPUT = ofv.GenerarInput("ORDESDE","TODOS","agregar",sForm,"30","dt","")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")







      Celda  = ofv.GenCelda("","tt","","","","","","PRIORIDAD HASTA","si")
      SINPUT = ofv.GenerarInput("ORHASTA","TODOS","agregar",sForm,"30","dt","")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")
      'Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")


      Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
	  sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
	  sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
	  Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
      formx    = FILA  & ofv.GenRow("","","","","40","","",Celda,"si")

      Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"_NEW","post","",formx,"si")


'-----------------------------









   sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","600 ","","0","","5","0","","",Formulario,"si")



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