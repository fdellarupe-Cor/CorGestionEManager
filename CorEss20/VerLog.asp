<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CsysLog.inc" -->
<%

ok=ofv.CheckUsuario()
'xfecLCID=Session.LCID
'Session.LCID=2057
if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  Server.ScriptTimeout=60*12
  APLCHECK()
  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
  OPC=REQUEST.QUERYSTRING("OPC")
  SEXPORT=REQUEST.FORM("EXPORT")
  APP=REQUEST.FORM("APP")
  set cn = ofv.conectar(ofv.strconn2)

  SELECT CASE SEXPORT

  CASE "XLS"

    Response.ContentType = "application/vnd.ms-excel"

  CASE "RFC3164"
     SET CLOG= NEW Syslog
     strsql="SELECT DESCAPLI FROM CORAPLIC WHERE CODAPLI='" & APP & "'"
    ' RESPONSE.WRITE STRSQL
	  set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  IF NOT RSX1.EOF THEN
	   SAPP=REPLACE(RSX1(0)," ","_")
	   SAPP=REPLACE(RSX1(0),"*","")
	  ELSE
	   SAPP=APP
	  END IF
	  CALL OFV.CERRARCONSULTA(RSX1)

     SFILE=SAPP & year(date())
     mes=month(date())
     if  mes > 10 then  mes="0" & month(date())

     dia=day(date())
     if  dia  < 10 then dia="0" & day(date())
     sfile=sfile & mes & dia & ".TXT"





 case else
    sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))


  END SELECT







SELECT CASE OPC

CASE "LST"
 '   if session("APLMDCLI")="BC" THEN
   STRSQL= "SELECT *,C.DESCAPLI,U.USRNAME FROM OBJALOGDET A INNER JOIN objalogext B ON (A.LOGID=B.USRLOGID) "
 '  else
 '  STRSQL= "SELECT *,C.DESCAPLI FROM OBJALOGDET A INNER JOIN objalogext B ON (A.LOGID=B.USRLOGID) "
 '  end if
   WSPSESION="SWHERE"
   SESSION(WSPSESION) =""
   SCONECTOR=""
   USR="TODOS" 'REQUEST.FORM("USR")
   FINICIO=REQUEST.FORM("FINICIO")
   FFIN=REQUEST.FORM("FFIN")
   TIPO=REQUEST.FORM("TIPO")
   APP=REQUEST.FORM("APP")

    IF FINICIO <> "TODOS" THEN

	       if IsDate(FINICIO) then
	          WWSORFECINIH = CLNG(CDATE(FINICIO)) - CLNG(CDATE("01-01-1900"))

	       else
	          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
	       end if

	       SORFECINIHW   = " (B.FOPER > = " & WWSORFECINIH & ") "
           SCONECTOR=" AND "
	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW


    END IF

    IF FFIN <> "TODOS" THEN

		       if IsDate(FFIN) then
		          WWSORFECINIH = CLNG(CDATE(FFIN)) -  CLNG(CDATE("01-01-1900"))

		       else
		          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
		       end if

		       SORFECINIHW   = " (B.FOPER  < = " & WWSORFECINIH & ") "

		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
		       SCONECTOR=" AND "

    END IF


   IF USR <> "TODOS" THEN


		       SORFECINIHW   = " (A.USUARIO ='" & USR & "') "

		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
		       SCONECTOR=" AND "

    END IF



    IF APP <> "TODOS" THEN


			       SORFECINIHW   = " (B.MASTEROBJ ='" & APP & "') "

			       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
			       SCONECTOR=" AND "

    END IF



 IF TIPO <> "TODOS" THEN


		       SORFECINIHW   = " (B.TOPER LIKE'%" & TIPO & "%') "

		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
		       SCONECTOR=" AND "
 else
              SORFECINIHW   = " (B.TOPER  in ('SKR-UPDATE','PAR-UPDATE','PAR-INSERT','PAR-DELETE','HABILITARUSR','SKR-INSERT','SKR-UPDATE','SKR-DELETE')) "

			   SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
		       SCONECTOR=" AND "


 END IF


   STRSQL =STRSQL & SESSION(WSPSESION)
  ' if session("APLMDCLI")="BC" THEN
   STRSQL = STRSQL &  " LEFT JOIN  (SELECT DISTINCT USRNAME,USUARIO AS USR1 FROM OBJALOG WHERE USRNAME IS NOT NULL) U ON(A.USUARIO=U.USR1)"
  ' end if
   STRSQL = STRSQL & "LEFT JOIN CORAPLIC C ON (B.MASTEROBJ=C.CODAPLI)"
   STRSQL = STRSQL & " ORDER BY FOPER,HOPER"
 ' response.write strsql
   set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   Celda = ofv.gencelda("","tt","","","","","","USUARIO","si")
   Celda = CELDA & ofv.gencelda("","tt","","","","","","NOMBRE","si")
   Celda = CELDA  & ofv.gencelda("","tt","","","","","","APLICACION","si")
   Celda = CELDA  & ofv.gencelda("","tt","","","","","","TIPO OPERACION","si")
   Celda = CELDA  & ofv.gencelda("","tt ALIGN=CENTER","","","","","","FECHA","si")
   Celda = CELDA  & ofv.gencelda("","tt ALIGN=CENTER","","","","","","HORA","si")
   Celda = celda & ofv.gencelda("","tt","","","","","","TABLA","si")
   Celda = celda & ofv.gencelda("","tt","","","","","","COLUMNA","si")
   Celda = celda & ofv.gencelda("","tt","","","","","","CODOBJ","si")
   Celda = celda & ofv.gencelda("","tt","","","","","","VALOR ANTERIOR","si")
   Celda = celda & ofv.gencelda("","tt","","","","","","VALOR ACTUAL","si")

   Fila =  ofv.genrow("","","","","","","",Celda,"si")
  SRFC3164=""
'  RESPONSE.WRITE STRSQL
  do while not rsx1.EOF
 ' session("SGU_nombre")=""
 '   XXX=GETINFOUSER(RSX1("USUARIO")&"")
  '  Celda = ofv.gencelda("","","","","","","",RSX1("USUARIO") & "-" & session("SGU_nombre"),"si")
   Celda = ofv.gencelda("","","","","","","",RSX1("USUARIO")  ,"si")

  ' if session("APLMDCLI")="BC" THEN
  '   IF RSX1("USRNAME") & "" ="" THEN
  '      XXX=GETINFOUSER(RSX1("USUARIO")&"")
  '        IF session("SGU_nombre")<> "NO" OR session("SGU_nombre") <> "" THEN
  '          STRSQL= " UPDATE OBJALOG SET USRNAME='"&session("SGU_nombre") & "' WHERE USUARIO='"& RSX1("USUARIO")&"'"
  '        '  RESPONSE.WRITE STRSQL
  '          set rsx1UPG = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  '         Celda = CELDA & ofv.gencelda("","","","","","","",session("SGU_nombre") ,"si")
  '        ELSE
  '        Celda = CELDA & ofv.gencelda("","","","","","","",session("SGU_nombre") ,"si")
  '       END IF
  '  ELSE
  ''  Celda = CELDA & ofv.gencelda("","","","","","","",RSX1("USRNAME") ,"si")
  '  END IF
  '  else
  '    XXX=GETINFOUSER(RSX1("USUARIO")&"")
  '   Celda = ofv.gencelda("","","","","","","",RSX1("USUARIO") ,"si")
  '   Celda = CELDA & ofv.gencelda("","","","","","","",session("SGU_nombre")  ,"si")
  '  end if

    Celda = CELDA & ofv.gencelda("","","","","","","",RSX1("USRNAME") ,"si")
    Celda = celda & ofv.gencelda("","","","","","","",RSX1("DESCAPLI") ,"si")
    Celda = celda & ofv.gencelda("","","","","","","",RSX1("TOPER") ,"si")
    Celda = celda & ofv.gencelda(""," ALIGN=CENTER	","","","","","",fechadorCor(CDATE(RSX1("FOPER")+  CLng(CDate("01-01-1900"))),""),"si")
     Celda = celda & ofv.gencelda("","AA ALIGN=RIGHT	","","","","","",GETGLBHMSW(RSX1("HOPER")),"si")
    Celda = celda & ofv.gencelda("","","","","","","",RSX1("TABLA"),"si")
    Celda = celda & ofv.gencelda("","","","","","","",RSX1("CLAVE"),"si")
    Celda = celda & ofv.gencelda("","","","","","","",RSX1("CODOBJ"),"si")
    Celda = celda & ofv.gencelda("","","","","","","",RSX1("VALANT"),"si")
    Celda = celda & ofv.gencelda("","","","","","","",RSX1("VALACT"),"si")
    Fila =  FILA & ofv.genrow("","","","","","","",Celda,"si")

    if SEXPORT="RFC3164" THEN
     DELIMXS=CHR(9)
     FECEVENTO=DELIMXS &fechadorCor(CDATE(RSX1("FOPER") +  CLng(CDate("01-01-1900"))),"") & DELIMXS
     HORAEVENTO=GETGLBHMSW(RSX1("HOPER")) & DELIMXS
     TIPO= RSX1("TOPER") & DELIMXS & DELIMXS & DELIMXS
     origen= RSX1("DESCAPLI") & delimxs
     USUARIO = RSX1("USUARIO") & DELIMXS
      IF RSX1("USRNAME") & "" ="" THEN
        XXX=GETINFOUSER(RSX1("USUARIO")&"")
        USRNAME=session("SGU_nombre")
      ELSE
        USRNAME=RSX1("USRNAME")
      END IF
      USRNAME=  USRNAME & DELIMXS
     IP=RSX1("DIRIP") & DELIMXS
     VALANT=RSX1("VALANT") & DELIMXS
     VALACT=RSX1("VALACT") & DELIMXS
     TABLA ="TABLA=" & RSX1("TABLA") & DELIMXS
     COLN="COLUMNA=" & RSX1("CLAVE")& DELIMXS
     CODOBJ="OBJID=" &RSX1("CODOBJ")

     MSG=FECEVENTO & HORAEVENTO & TIPO & origen & USUARIO & USRNAME & IP & VALANT & VALACT & TABLA &COLN & CODOBJ


     SRFC3164=SRFC3164 & CLOG.SaveMsg(1,6,CDATE(RSX1("FOPER") +  CLng(CDate("01-01-1900"))),GETGLBHMSW(RSX1("HOPER")),RSX1("DIRIP"),MSG,XEXTRA)
    END IF
    rsx1.movenext

   LOOP
   CALL OFV.CERRARCONSULTA(RSX1)
   FORMULARIO=FILA
CASE ELSE
    sForm="FORM1"

    sInput = "FECHA INICIO"
    Celda = ofv.gencelda("","TTT","","","","","",sInput,"si")

    sInput = ofv.GenerarInput("FINICIO",date(),"text",sForm,"50","dt","no")
    Celda = CELDA & ofv.gencelda("","","","","","","",sInput,"si")
    Fila =  ofv.genrow("","","","","","","",Celda,"si")


     Celda = ofv.gencelda("","ttT","","","","","","FECHA FIN","si")

     sInput = ofv.GenerarInput("FFIN",date(),"text",sForm,"50","dt","no")
     Celda = Celda & ofv.gencelda("","","","80%","","","",sInput,"si")
     Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


'     Celda = ofv.gencelda("","ttT","","20%","","","","USUARIO","si")
 '    strsql="select distinct usuario from OBJALOGDET"
  '   set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   '  scomK="TODOS"
    ' scomD="TODOS"
'     SCOMSEP = chr(9)
''     do while not rsx1.EOF
'	  scomK =scomK & SCOMSEP & RSX1(0)
'	  XXX=GETINFOUSER(RSX1(0) &"")
'	  scomD = scomD & SCOMSEP & session("SGU_nombre")
'	  RSX1.MOVENEXT
'	 LOOP
'	 CALL OFV.CERRARCONSULTA(RSX1)
'	 sInput = ofv.gencomboFX(sForm,"USR","TODOS",scomK,scomD,"no")
'	 Celda = Celda & ofv.gencelda("","","","80%","","","",sInput,"si")




'     Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda = ofv.gencelda("","ttT","","20%","","","","PRODUCTOS","si")
	  strsql="SELECT CODAPLI,DESCAPLI FROM CORAPLIC"
	  set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  scomK="TODOS"
	  scomD="TODOS"
	  SCOMSEP = chr(9)
	  do while not rsx1.EOF
	   scomK =scomK & SCOMSEP & RSX1(0)
	   scomD = scomD & SCOMSEP & RSX1(1)
	   RSX1.MOVENEXT
	  loop
	 CALL OFV.CERRARCONSULTA(RSX1)
	 sInput = ofv.gencomboFX(sForm,"APP","TODOS",scomK,scomD,"no")
	 Celda = Celda & ofv.gencelda("","","","80%","","","",sInput,"si")

     Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda = ofv.gencelda("","ttT","","20%","","","","TIPO","si")


	 scomK = "TODOS" & chr(9) &"UPDATE" & chr(9) & "INSERT" & chr(9) &"DELETE"
	 scomD="TODOS" & CHR(9) & "UPDATE" & CHR(9) & "INSERT"  & CHR(9) & "DELETE"
     sInput = ofv.gencomboFX(sForm,"TIPO","TODOS",scomK,scomD,"no")
     Celda = Celda & ofv.gencelda("","","","80%","","","",sInput,"si")

     Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda = ofv.gencelda("","ttT","","20%","","","","Formato Exportacion","si")
     SCOMD= "XLS" & CHR(9) &"RFC3164" & CHR(9) & "NO"
	 scomK=SCOMD
	 sInput = ofv.gencomboFX(sForm,"EXPORT","NO",scomK,scomD,"no")
	 Celda = Celda & ofv.gencelda("","","","80%","","","",sInput,"si")

     Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     SINPUT = "<BR><INPUT TYPE='SUBMIT' VALUE='SELECCIONAR' CLASS='BT'>&nbsp;&nbsp;"
     Celda = ofv.gencelda("","ut ","LEFT","","","","",SINPUT,"si")
     Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
'''''''''''''''''''''''''

     SACCION = "VERLOG.ASP?OPC=LST"


   Formulario = ofv.genform(sForm,"",sAccion,"","post TARGET=_NEW","",Fila,"si")

    END SELECT

    sHTML = sHTML & ofv.GenTabla("","aaa","","","","0","","0","0","","",Formulario,"si")

  'Botonera
   Agregar  = ""
   Buscar   = ""
   Retornar = ""


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Retornar)


  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</body></html>"




SELECT CASE SEXPORT

CASE  "XLS"
 Response.Write ofv.GenTabla("","aa BORDER","","80%","","0","","0","0","","",Formulario,"si")
CASE "RFC3164"
 Response.ContentType = "application/vnd.text"
 Response.AddHeader "content-disposition", "attachment;filename=" & sfile
 Response.Write sRFC3164

case else
 response.write shtml
END SELECT


end if
'Session.LCID=xfecLCID
%>