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


  proceso    = request.querystring("proceso") & ""
  subproceso = request.querystring("suBproceso") & ""


  'proceso=clng("0" & request.FORM("proc"))
 ' response.write proceso
 ' subproceso=request.FORM("suBproceso")&""

  if subproceso="0" OR subproceso="TODOS"  OR subproceso="nsnc" or proceso=0 then
   Iraiz=proceso
  else
   Iraiz=subproceso
  end if

  strsql =  "select   COALESCE(e.descri,D.DESCRIACC) as actividad , A.*  ,d.secuen  "
  strsql =  strsql &" from csrelpathdescri a  "
  strsql =  strsql &" inner join objetos b on (a.cobj = b.codobjeto)  "
  Strsql =  Strsql & "  inner join objprty n on (b.codobjeto=n.codobjeto and estado=3)"
  strsql =  strsql &"      inner join objlnk x on (a.cobj = x.codobjeto and (x.ltipo = 'LRGVN' or x.ltipo = 'LRGCR'))    "
  strsql =  strsql &"        inner join objproced c "
  strsql =  strsql &"   on ( c.codobjeto = x.codobjetol and c.codproced = x.iden4 )  "
  strsql =  strsql &"      inner join objaccion d on ( C.codproced = d.codproced) "
  strsql =  strsql &"          LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt and x.ltipo = 'LRGVN' "
  strsql =  strsql &"          )"
  strsql =  strsql &"          where      ( "
  strsql =  strsql &"          A.MC1 =" & Iraiz & " OR  A.MC2 = " & Iraiz & " OR  A.MC3 = " & Iraiz
  strsql =  strsql &"          OR  A.MC4 = " & Iraiz & " OR  A.MC5 = " & Iraiz & " OR  A.MC6 = " & Iraiz & " OR  "
  strsql =  strsql &"          A.MC7 = " & Iraiz & " OR  A.COBJ =" & Iraiz & "   ) "
  strsql =  strsql &"          AND (  B.APLIC = 'RG' ) "
  strsql =  strsql &"           ORDER BY  A.TSEC,D.SECUEN"


'   strsql =  strsql &"          where      (   A.MC6 = "& Iraiz & " OR  A.MC7 = "& Iraiz & " OR  A.COBJ ="& Iraiz & "   )  ORDER BY  A.TSEC,D.SECUEN"
'''RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

   xfila=""
'  Celda = Celda & ofv.gencelda("","UTI style='background-color:#8888ff;color:#ffffff' " ,"","40%","","","","PROCESO","si")
  if not rs.eof then
   sProceso=rs("expr6")
  end if

  Celda =  ofv.gencelda("","UTI style='background-color:#8888ff;color:#ffffff' COLSPAN=4","","1%","","","","PROCESO :&nbsp;" & SPROCESO,"si")

  xFila  = xFILA & ofv.genrow("","aaa bgcolor='#f9f9f9' ","","","","","",Celda,"si")
  Celda =  ofv.gencelda(""," COLSPAN=4","","","","","","&nbsp;","si")
  xFila  = xFILA & ofv.genrow(""," ","","","","","",Celda,"si")
  Celda = ofv.gencelda("","uti  " ,"","","","","","<b>Nº</b>","si")
  Celda = Celda & ofv.gencelda("","UTI " ,"","20%","","","","<b>SUBPROCESO<b/>","si")
  Celda = Celda & ofv.gencelda("","UTI " ,"","","","","","<b>Nº</b>","si")
  Celda = Celda & ofv.gencelda("","UTI " ,"","","","","","<b>ACTIVIDADES</b>","si")

  xFila  = xFILA & ofv.genrow("","aaa bgcolor='#F7EEDD' ","","","","","",Celda,"si")

  ISubProceso=1
  IActividad=1
  Ssubproceso=""
  ifila=0
  fila=""
          xTrColor = "#EEEEEE"
  do while not rs.eof


    if ifila=30 then
       sHTML=shtml & xshtml &  ofv.GenTabla("","AAA   " & wkstylo1 ,"","600 ","","0","","0","-1","","",xfila & Fila,"si") & "<DIV >&nbsp;</DIV> "
      ' sHTML=shtml & xsHTML&  "<DIV style='page-break-before :always'></DIV> "
      ifila=0
      fila=""
      Celda = ""
    end if

     if  Ssubproceso <> rs("expr7") then
        Celda =  ofv.gencelda(""," COLSPAN=4 ","","","","","","&nbsp;","si")
        Fila  = FILA & ofv.genrow(""," ","","","","","",Celda,"si")
	    Ssubproceso=rs("expr7")
	    Celda=ofv.gencelda("","uti bgcolor=#EEEEEE" ,"","","","","",ISubProceso,"si")
	    Celda = celda & ofv.gencelda("","uti  bgcolor=#EEEEEE" ,"","","","","","<b>" & UCASE(Ssubproceso &"") &"</b>","si")
	    ISubProceso=ISubProceso+1
	    IActividad=1
	 else
	   Celda = ofv.gencelda("","uti  " ,"","","","","","&nbsp;","si") & ofv.gencelda("","uti" ,"","","","","","&nbsp;","si")
     end if


     Celda = Celda & ofv.gencelda("","uti  " ,"","","","","",IActividad,"si")
     Celda = Celda & ofv.gencelda("","uti  " ,"","","","","",UCASE(rs("actividad")&""),"si")
     Fila  = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

       'if xTrColor = "#FFFFFF" then
          xTrColor = "#EEEEEE"
     '  ELSE
     '     xTrColor = "#FFFFFF"
     '  END IF



     IActividad=IActividad+1
     rs.movenext
     ifila=ifila +1
  loop

  if fila<>"" then
   sHTML=shtml & xshtml &  ofv.GenTabla("","AAA  " & wkstylo1 ,"","600 ","","0","","0","-1","","",xfila & Fila,"si") '& "<DIV style='page-break-before:always'></DIV> "
   ifila=0
    fila=""
    Celda = ""
  end if
  call ofv.cerrarconsulta(rs)



  'sHTML=sHTML & ofv.GenTabla("","AAA  " & wkstylo1 ,"","100% ","","0","","0","-1","","",Fila,"si")
case else


    XTPPER      = "" & request.form("PERIODO")
    XTPROC      = clng("0" & request.form("PROC"))
    XSUBPROCESO = "" & request.form("SUBPROCESO")

  '  response.write request.form


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

      scodmaster = WKORPERIODOMAP
      stipmaster = "PROCESO"




      Celda  = ofv.GenCelda("","tt","","","","","","PROCESO","si")

      Strsql = "Select OBJETOS.* From relaobj, OBJETOS  where CODOJASOC = CODOBJETO "
      Strsql = Strsql & " AND codobjmaster = '" & scodmaster & "' and tipobjmaster = '" & stipmaster & "'  "
      Strsql = Strsql & " ORDER BY OBJETOS.INDOBJ "
    '  RESPONSE.WRITE STRSQL

      IF XTPROC = 0 THEN

         set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if NOT rsX1.eof then
            XTPROC          = RSX1(0)

         END IF

      ELSE 


      END IF

      sInput = ofv.GenerarCombo("Subconsulta0","PROC",XTPROC,Strsql,cn,0,3,"")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

           Strsql ="SELECT   ax.cobj,B.DESCRI , O1.descri, O2.descri AS Expr1, O3.descri AS Expr2, O4.descri AS Expr3, O5.descri AS Expr4, "
	   Strsql = strsql & "  O6.descri AS Expr5, O7.descri AS Expr6,"
	   Strsql = strsql & "  O8.descri AS Expr7, ax.mc1, ax.mc2, ax.mc3, ax.mc4, ax.mc5, ax.mc6, ax.mc7, ax.cobj, ax.TSEC "
	   Strsql = Strsql & " FROM         dbo.CsRelPath ax LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O1 ON ax.mc1 = O1.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O2 ON ax.mc2 = O2.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O3 ON ax.mc3 = O3.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O4 ON ax.mc4 = O4.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O5 ON ax.mc5 = O5.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O6 ON ax.mc6 = O6.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O7 ON ax.mc7 = O7.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                dbo.OBJETOS O8 ON ax.cobj = O8.codobjeto "
	   Strsql = Strsql & " inner join objetos b on (ax.cobj = b.codobjeto) "
	   Strsql = Strsql & "  inner join objprty n on (ax.cobj=n.codobjeto and estado=3)"
	   Strsql = Strsql & " WHERE (B.CODVINOBJ = 'VN' or B.CODVINOBJ = 'CR' or B.CODVINOBJ = 'AI') AND ("
	   Strsql = Strsql & " Ax.MC1 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC2 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC3 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC4 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC5 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC6 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC7 = 0" & XTPROC & "  "
	   Strsql = Strsql & "  )"

	   '   response.write "<br>" & strsql & "<br>" & chr(13)
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       scomK="TODOS" & CHR(9)
       scomD="TODOS" & CHR(9)
       do while not rs.eof
        scomK = scomK & rs(0) & chr(9)
        scomD = scomd & rs(1) & chr(9)
        rs.movenext
       loop

       call ofv.cerrarconsulta(rs)

       IF XSUBPROCESO = "" THEN XSUBPROCESO = "TODOS" 

       sInput = ofv.gencomboFX("Subconsulta0","SUBPROCESO",XSUBPROCESO,scomK,scomD,"")

      Celda  = ofv.GenCelda("","tt","","","","","","SUBPROCESO","si")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")
      sAccion    = "ormapaproceso.asp?Opc=SET"
      Formulario = Formulario & ofv.GenForm("Subconsulta0","",sAccion,"","post","",Fila,"si")


      Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
      sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
      sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
      Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
      formx    = ofv.GenRow("","","","","40","","",Celda,"si")
      sAccion    = "ormapaproceso.asp?Opc=LST&proceso=" & XTPROC &"&subproceso=" & XSUBPROCESO
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