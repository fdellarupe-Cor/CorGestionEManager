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

  sOPCION       = request.querystring("OPCION")
  SCODINDI       = request.querystring("INDI")
  SACTIVIDAD=    request.querystring("ACTIVIDAD")
  SRIESGO=    request.querystring("RIESGO")
  scodmaster = request.querystring("codmaster") & ""
  stipmaster = request.querystring("tipmaster") & ""
  idindicador=request.querystring("idindicador")
  vuelta2 = request.querystring("volver")
  ' vuelta        = ofv.vueltaasp2("../")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_VOLVER="&vuelta2&"_OPCION=LISTAR_RIESGO=" & SRIESGO & "_ACTIVIDAD=" & SACTIVIDAD & "_CODMASTER=" &SCODMASTER & "_TIPMASTER=" & stipmaster
 ' response.write request.querystring
  'RESPONSE.WRITE SRIESGO
  sHTML = ofv.FormHeader(session("FormN"))
   'sHTML = ofv.FormHeader("")
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  IF SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4   AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 )  THEN

  bAgregar = cbool(1)
  bModificar = cbool(1)
  bSuprimir = cbool(1)

  else

  bAgregar = cbool(0)
  bModificar = cbool(0)
  bSuprimir = cbool(0)

  end if


  set cn = ofv.conectar(ofv.strconn2)


   SESSION("XWIN")=""
   SESSION("ADDRELIND")=""
'''''''''''''''''''''''''''''''''''''''

if sopcion = "ASGINDI" then

        sfechai = clng(date()) - clng(cdate("01-01-1900"))


        STRSQL="SELECT  * FROM ORRELAINDIRIESGO WHERE IDINDICADOR=0"& idindicador & " AND CODRIESGO='" & SRIESGO &"'"
        'RESPONSE.WRITE STRSQL
        SET RSREL =ofv.crearconsultaEx(StrSql,cn,1,volver)

        IF RSREL.EOF THEN
         strsql = "Insert into ORRELAINDIRIESGO "
         strsql = strsql & " (FECHA,USUARIO,CODRIESGO"
         strsql = strsql & " ) Values ("
         strsql = strsql &   sfechai
         strsql = strsql & ",'" & session("usuario") & "','"&SRIESGO&"')"
         'response.write strsql
         call ofv.crearconsultaEx(StrSql,cn,1,volver)
         SOPCION="LISTAR"
        ELSE
         SOPCION="EXISTEREL"
        END IF

        CALL OFV.CERRARCONSULTA(RSREL)

  end if



'' INCORPORO INDICADOR

  if sopcion = "SIA" then

        sfechai = clng(date()) - clng(cdate("01-01-1900"))

        strsql = "Insert into ORINDICADORES "
        strsql = strsql & " (acceso,lstusr,FORMULA,CANTVAR,TIPO,DESCRIPCION,FECHACREACION,estado,usuario"
        strsql = strsql & " ) Values (1,1,16,1,0,"
        strsql = strsql &  "'NUEVO INDICADOR'," & sfechai
        strsql = strsql & ",1,'" & session("usuario") & "')"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)





  end if


  if sOPCION = "ELIM" then

	strsql = "DELETE FROM ORRELAINDIRIESGO WHERE OBJID = 0" & SCODINDI
	call ofv.crearconsultaEx(StrSql,cn,1,volver)
	SOPCION="LISTAR"

  END IF

'--------------------
   celda=""
   Celda = Celda & ofv.GenCelda("","uti style='background-color:#888888;color:#ffffff;' ","","10%","","","","PERIODO:","si")
   Celda = Celda & ofv.GenCelda("","uti ","","65%","","","",session("ORPERIODODESC"),"si")
   Celda = Celda & ofv.GenCelda("","ut  ","","20%","","","","","si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   wktabperiodo = ofv.GenTabla("","","","98% ","","0","","-1","1","","",Fila,"si")

    sImagePRC =""' "<font face='corflaticon' size=2 color='" & wspTIPOCLC & "'  >&#060;</font>"

    WKDIVSECUEN = ""'GENPROCSECUENEX()

    WKPROCDESC =""' WKDIVSECUEN & sInput & ":www&nbsp;" & WKPROCDESC
    sImage = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0233;</font>"

    wkproccod =  vuelta2 &"?" & request.querystring ' "ORCIRCUITO.ASP?ACTIVIDAD=" & SACTIVIDAD & "&CONTROLTIPO=" & SCONTROLTIPO
    wkproccod = "<a class=nt  href='" & wkproccod & "' ><b>"
    wkproccod = wkproccod & sImage
    wkproccod = wkproccod & "</b></a>"


    wkproccodN = "&nbsp;"

     'agregar relacion indicador

    if bModificar then
       WKSPTITLE="Vincular indicador"
       sImage     = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"
       saccion    = "orasignaindicador.asp?opcion=ASGINDI&RIESGO=" & SRIESGO & "&ACTIVIDAD=" & SACTIVIDAD
       saccion    = saccion & "&CODMASTER=" &SCODMASTER & "&TIPMASTER=" & stipmaster & "&volver=" & vuelta2
       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='" & WKSPTITLE & "'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"
    end if

     STRSQL="SELECT * FROM CSORBUSQRIESGO WHERE CODRIESGO='" & SRIESGO & "'"
  	 SET RSRIESGOPATH=ofv.crearconsultaEx(StrSql,cn,1,parametros)
  	 IF NOT RSRIESGOPATH.EOF THEN
  	  SINPUT= RSRIESGOPATH("DESCRIACT") & " / " & RSRIESGOPATH("RIESGONOMBRE")
  	 ELSE
  	 SINPUT="&nbsp;"
  	 END IF
  	 CALL OFV.CERRARCONSULTA(RSRIESGOPATH)

     celda=ofv.GenCelda("","Ut style='background-color:#f0f0f0;' ","","90%","","","",SINPUT,"si")
     Celda = Celda & ofv.GenCelda("","Ut style='background-color:#f0f0f0;color:#990000;' ","CENTER","1%","","","",sImagePRC,"si")
     Celda = Celda & ofv.GenCelda("","ut style='background-color:#f0f0f0;color:#888888;' ","","85%","","","", WKPROCDESC,"si")
     Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","CENTER","1%","","","",wkproccodN,"si")
     'Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","CENTER","1%","","","",wkproccodN,"si")
     Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","CENTER","1%","","","",wkproccod,"si")
     Fila  = ofv.GenRow("","","","","","","",Celda,"si")
     TablaH = ofv.GenTabla("","aaa style='background-color:#f0eedd;border-style:solid;border-width:1;border-color:#000000;' ","","98% ","","0","","-1","1","","",Fila,"si")
     celda=""
     Fila=""
'----------------------------------------------------------------
     IF SOPCION <> "CONSULTA" THEN shtml=shtml &  wktabperiodo & TablaH

  SELECT  case Sopcion


  case "CONSULTA"

	    Celda = ""

	       wkproccod = "&nbsp;"


	       Formulario = ""

	        Celda = ""

	     '  STRSQL="SELECT a.objid,a.idindicador,B.* ,(select c.descri from orfrecuencia c where c.objid=b.unidadfrecuen) as frecuencia  FROM ORRELAINDIRIESGO A,ORINDICADORES B WHERE A.CODRIESGO='" & SRIESGO & "'"
	     '  STRSQL = STRSQL & " AND  A.IDINDICADOR*=B.OBJID	"

	          STRSQL="SELECT a.objid,a.idindicador,B.* ,(select c.descri from orfrecuencia c where c.objid=b.unidadfrecuen) as frecuencia  FROM ORRELAINDIRIESGO A "
	          strsql=  strsql & " left join ORINDICADORES b on (a.idindicador=b.objid) "
	          strsql = strsql & " WHERE A.CODRIESGO='" & SRIESGO & "'"

	          'ORINDICADORES B WHERE A.CODRIESGO='" & SRIESGO & "'"
		   '  STRSQL = STRSQL & " AND  A.IDINDICADOR*=B.OBJID	"


	   '   response.write strsql
	      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	      if rs.eof then

	        Celda = ofv.gencelda("","vT  style='border-style:SOLID;border-width:1;border-color:#f93333;' ","center","","","","","<big>NO HAY ASIGNADOS INDICADORES</big>","si")
	        Formulario = ofv.genrow("","","","","60","","",Celda,"si")

	      ELSE

	        Celda = ""
	       ' Celda = ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","1%","","","","","si")
	        Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","NOMBRE&nbsp;DEL&nbsp;INDICADOR","si")
	        'Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","DESCRIPCION","si")


	        Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","DESCRIPCION","si")
	        Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","RESPONSABLE","si")
	        Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","FRECUENCIA","si")
	        'Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","VER UMBRALES","si")
	        Formulario = ofv.Encabezados("Descripcion","fecha","formula","cant var")

	        Formulario = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

	      posI = 0

	      WKSETTABLA="ORRELAINDIRIESGO"

	      Do While (Not rs.EOF)
	        sForm = "form" & posi

	        modiF=Fmodifica(rs("estado"),"","")
	        Celda = ""

            ' Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")






	        'STRSQL="SELECT * FROM orindicadores  where estado=2 ORDER BY DESCRIpcion"
	 	    sInput =rs("descripcion")'"" ofv.GenerarCombo(sForm,rs(1).name & "READONLY" ,rs(1).value,Strsql,cn,0,1,"")


	        'Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
	        Celda = Celda & ofv.gencelda("","ut style='border-style:OUTSET;border-width:1;'   align=left","","","","","", sInput,"si")

	        'descripcion


	    	       nCol = 60


	               wdescri =""
	    	       if  not  isnull(rs(5)) then
	    	           wdescri = rs(5)
	    	       end if

	    	       if  len(wdescri) = 0 then wdescri = " "

	    	       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' "
	    	       t3b = t3b & "name=" & rs(5).name & " cols=" & nCol & " rows=2 "
	    	       if  bModificar and modiF then
	    	           t3b = t3b & " onblur=" & sForm & ".submit() "
	    	           t3b = t3b & " onkeypress=" & chr(34)
	    	           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
	    	           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
	    	       else
	    	           t3b = t3b & " readonly >"
	    	       end if
	    	       xdatos =  wdescri
	    	       t3b    = t3b & xdatos & "</textarea></font>"
	    	       SINPUT = wdescri 't3b



	         Celda = Celda & ofv.gencelda("","UT style='border-style:OUTSET;border-width:1;' align=left","","","","","",sInput,"si")



	        CALL GETINFOUSER(RS("USUARIO") )

	        SINPUT=session("SGU_nombre")
	        Celda = Celda & ofv.gencelda("","ut style='border-style:OUTSET;border-width:1;'  align=left","","","","","",sInput,"si")

	          SINPUT = rs("frecuencia")
	 	      Celda = Celda & ofv.gencelda("","ut style='border-style:OUTSET;border-width:1;'  align=left","","","","","",sInput,"si")


	       Filax  = ofv.genrow("","AAA  bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

	        sAccion = ""
	        if bModificar then
	          sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
	          sAccion = sAccion & "&tabla=ORrelaindiriesgo&volver=" & vuelta
	        end if
	        Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")


	        posi = posi + 1
	        rs.Movenext
	      Loop
	      END IF
     sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")


  case "EXISTEREL"

       Celda = ofv.gencelda("","vT  style='border-style:SOLID;border-width:1;border-color:#f93333;' ","center","","","","","<big>ESTA ASIGNADO ESE INDICADOR</big>","si")
       Formulario = ofv.genrow("","","","","60","","",Celda,"si")
       sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")

 case "LISTAR"

   Celda = ""

      wkproccod = "&nbsp;"


      Formulario = ""

       Celda = ""

      STRSQL="SELECT a.objid,a.idindicador,B.* ,(select c.descri from orfrecuencia c where c.objid=b.unidadfrecuen) as frecuencia  FROM ORRELAINDIRIESGO A "
      strsql=  strsql & " left join ORINDICADORES b on (a.idindicador=b.objid) "
      strsql=  strsql & " WHERE A.CODRIESGO='" & SRIESGO & "'"
      'STRSQL = STRSQL & " AND  A.IDINDICADOR*=B.OBJID	"
     'response.write strsql
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rs.eof then

       Celda = ofv.gencelda("","vT  style='border-style:SOLID;border-width:1;border-color:#f93333;' ","center","","","","","<big>NO HAY ASIGNADOS INDICADORES</big>","si")
       Formulario = ofv.genrow("","","","","60","","",Celda,"si")

     ELSE

       Celda = ""
       Celda = ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","1%","","","","","si")
       Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","NOMBRE&nbsp;DEL&nbsp;INDICADOR","si")
       'Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","DESCRIPCION","si")


       Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","DESCRIPCION","si")
       Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","FORMULA","si")
       Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","RESPONSABLE","si")
       Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","FRECUENCIA","si")

       Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","RESULTADOS","si")
       Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","VER UMBRALES","si")
       Formulario = ofv.Encabezados("Descripcion","fecha","formula","cant var")

       Formulario = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

     posI = 0
     sImageELN    = "<font face='corflaticon' size=3 color='#dddddd'>&#0253;</font>"
     sImageELS    = "<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
     WKSETTABLA="ORRELAINDIRIESGO"
   '  response.write strsql
     Do While (Not rs.EOF)
       sForm = "form" & posi

       modiF=Fmodifica(rs("estado"),"","")
       Celda = ""

       'Eliminar
   	    if bSuprimir   then

   	      wktitdelete = "Confirma la Desvinculacion ?"
   	      xsAccion="ORASIGNAINDICADOR.asp?opcion=ELIM&INDI=" & RS(0) & "&RIESGO=" & SRIESGO & "&ACTIVIDAD=" & SACTIVIDAD & "&CODMASTER=" &SCODMASTER & "&TIPMASTER=" & stipmaster & "&volver=" & vuelta2

          sImageEL = sImageELS
          sformobj = "FRMRGA" & posi
          formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",xsAccion,"","post","","","si")
          sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "
   	      xCelda = ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ","","2%","","","",sImageEL,"si") & formobj '""'ofv.BotonEliminar(rs(0),sHref)

   	    else
   	      xCelda = ofv.GenCelda("","","","","","","","&nbsp;","si")
   	    end if







       STRSQL="SELECT * FROM orindicadores  where estado=2 ORDER BY DESCRIpcion"
	   sInput = ofv.GenerarCombo(sForm,rs(1).name,rs(1).value,Strsql,cn,0,1,"")


       'Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
       Celda = Celda & ofv.gencelda("","AAA ALIGN=LEFT","","","","","", sInput,"si")

       'descripcion


   	       nCol = 60


              wdescri =""
   	       if  not  isnull(rs(5)) then
   	           wdescri = rs(5)
   	       end if

   	       if  len(wdescri) = 0 then wdescri = " "

   	       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' "
   	       t3b = t3b & "name=" & rs(5).name & " cols=" & nCol & " rows=2 "
   	       if  bModificar and modiF then
   	           t3b = t3b & " onblur=" & sForm & ".submit() "
   	           t3b = t3b & " onkeypress=" & chr(34)
   	           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
   	           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
   	       else
   	           t3b = t3b & " readonly >"
   	       end if
   	       xdatos =  wdescri
   	       t3b    = t3b & xdatos & "</textarea></font>"
   	       SINPUT = wdescri 't3b



        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")

        STRSQL="SELECT * FROM CSORFORMULAS  where OBJID=0" & RS("FORMULA")
        set rsF = ofv.crearconsultaEx(StrSql,cn,1,parametros)

        IF NOT RSF.EOF THEN
           select case RSF("DEF")
            case "PORCENTAJE"
              SINPUT=RSF("DEF") & "&nbsp;(A/B)*100"
            case "SUMA"
               SINPUT=RSF("DEF") & "&nbsp;" & rs("cantvar") & "&nbsp;VARIABLES"
            case "UNICO"
               SINPUT="VALOR&nbsp;"& RSF("DEF")
           end select
        else
          SINPUT=""
        END IF
        CALL OFV.CERRARCONSULTA(RSF)
        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")
       IF (RS("USUARIO")&"") <> "" THEN
        CALL GETINFOUSER(RS("USUARIO") )
        SINPUT=session("SGU_nombre")
       ELSE
        SINPUT="NO"
       END IF
       Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")

         SINPUT = rs("frecuencia")
	     Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")

       if (rs(1) & "") = "" then

          Celda = Celda & ofv.gencelda("","ut","center","","","","","---","si")
          Celda = Celda & ofv.gencelda("","ut","center","","","","","---","si")

       else

          sImagenP = "&#0232;"
	  sFuncAnc = ""
	  sinput = "<a title='Ver Resultados.' ONCLICK=ALERT('1'); class='ft1n' " & sFuncAnc & " href='ORCARGAINDICADORES.asp?OPCION=SI&ID=" & rs(0) & "&CantRegMover=" & CantRegMoverx
	  sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx  & "' >" & sImagenP & "</a>"

	  sImageOF   = "<font face='corflaticon' size=2 color='#FF0000'  >" & sImagenP & "</font>"
	  url="'orcargavariables.asp?opcion=VERINFINDI&id=" &rs(1) & "' ,'mywindow','status=0,toolbar=0,width=700,height=400'"

          Celda = Celda & ofv.gencelda("","","CENTER title='Ver Resultados.' onclick="&chr(34)&"window.open("&url&");"&chr(34)  & " style='cursor:pointer;'","","","","",sImageOF,"si")

          sImagenP = "&#0232;"
	   sFuncAnc = ""
	   sinput = "<a title='Ver Resultados.'  class='ft1n' " & sFuncAnc & " href='ORCARGAINDICADORES.asp?OPCION=SI&ID=" & rs(0) & "&CantRegMover=" & CantRegMoverx
	   sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx  & "' >" & sImagenP & "</a>"

	   sImageOF   = "<font face='corflaticon' size=2 color='#00FF00'  >" & sImagenP & "</font>"
	   url="'ORVERUMBRALES.asp?opcion=VERINFINDI&id=" &rs(1) & "' ,'mywindow','status=0,toolbar=0,width=700,height=400'"

          Celda = Celda & ofv.gencelda("","","CENTER title='Ver Umbrales.' onclick="&chr(34)&"window.open("&url&");"&chr(34)  & " style='cursor:pointer;'","","","","",sImageOF,"si")

       end if


       sAccion = ""
	   if bModificar then
		sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
		sAccion = sAccion & "&tabla=ORrelaindiriesgo&volver=" & vuelta
       end if

       Filaxs  = ofv.GenForm(sForm,"",sAccion,"","post","",celda,"si") '
       Filax=  ofv.genrow("","AAA   bgcolor=" & xTrColor & " ","","","","","",xCelda & Filaxs,"si")
       Formulario = Formulario & Filax

     '  Formulario = Filax 'Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")


       posi = posi + 1
       rs.Movenext
     Loop
     END IF
  sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")


  CASE ELSE



   Strsql = "select * FROM orindicadores WHERE  ESTADO=2 AND tipo<>999"

   Celda = ""

   wkproccod = "&nbsp;"

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","TABLA DE INDICADORES","si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")




  'Encabezados
  Formulario = ""

    Celda = ""


    Formulario = ""



' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

    Celda = ofv.gencelda("","vT  style='border-style:SOLID;border-width:1;border-color:#f93333;' ","center","","","","","<big>NO HAY CARGADOS INDICADORES</big>","si")
    Formulario = ofv.genrow("","","","","60","","",Celda,"si")

  ELSE

    Celda = ""
    Celda = ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","1%","","","","","si")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","NOMBRE&nbsp;DEL&nbsp;INDICADOR","si")



    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","FECHA CREACION","si")
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","CREADOR","si")
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","ASIGNAR","si")
    Formulario = ofv.Encabezados("Descripcion","fecha","formula","cant var")
    Formulario = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  posI = 0

  WKSETTABLA="orindicadores"


  Do While (Not rs.EOF)
    sForm = "form" & posi

    modiF=Fmodifica(rs("estado"),"","")
    Celda = ""

    'Eliminar
	    if bSuprimir  and modiF then
	      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
	      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
	      Celda = ofv.BotonEliminar(rs(1),sHref)
	    else
	      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
	    end if


    'ID

    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'indicador
    if modiF then
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"70","dt","")
    else
     sinput=rs(1)
    end if
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")




    sInput = GTFechadorSH(RS("FECHACREACION"))
	Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")


     IF (RS("USUARIO")&"") <> "" THEN
	  CALL GETINFOUSER(RS("USUARIO") )
	  SINPUT=session("SGU_nombre")
	 ELSE
	  SINPUT="NO"
     END IF
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    STrsql="SELECT  * FROM ORRELAINDIRIESGO WHERE IDINDICADOR="& rs(0) & " AND CODRIESGO='" & SRIESGO &"'"
    set rsrel= ofv.crearconsultaEx(StrSql,cn,1,parametros)

    if rsrel.eof then
     sImagenP = "&#0232;"
  	 sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
  	 sinput = "<a class='ft1n' " & sFuncAnc & " href='orasignaindicador.asp?idindicador="&rs(0)&"&opcion=ASGINDI&RIESGO=" & SRIESGO & "&ACTIVIDAD=" & SACTIVIDAD & "&CODMASTER=" &SCODMASTER & "&TIPMASTER=" & stipmaster & "&volver=" & vuelta2
  	 sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
  	 sinput = sinput  & "' title='ASIGNAR " & rs(1) & "' >" & sImagenP & "</a>"
  	else
  	  sinput=""
  	end if
    Celda = Celda & ofv.gencelda("","ft","LEFT","","","","",sInput,"si")

   call ofv.cerrarconsulta(rsrel)


    Filax  = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=ORINDICADORES&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")


    posi = posi + 1
    rs.Movenext
  Loop
  END IF
  sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")


 end select
  'Botonera
  'Botonera
   Agregar = ""
   if bAgregar  then
     sAccion = "ORindicadores.asp?XWIN=ADDRELIND&RIESGO=" & SRIESGO & "&ACTIVIDAD=" & SACTIVIDAD & "&CODMASTER=" &SCODMASTER & "&TIPMASTER=" & stipmaster & "&volver=" & vuelta2
     Agregar = ofv.BotonAgregar(sAccion)
   end if
   Buscar  = ""

  sAccion = vuelta2 &"?" & request.querystring
  Celda = ofv.BotonVolver(sAccion)


   if Sopcion ="CONSULTA" then  AGREGAR="no"
  'sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","no","","")
  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","",AGREGAR,"","no","")

  call ofv.cerrarconsulta(rs)




  call ofv.cerrarconn(cn)
  sHTML = sHTML & "<script type='text/javascript'>"
      sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
      sHTML = sHTML & "  {"
      sHTML = sHTML & "  var r=confirm(disp_msg);"
      sHTML = sHTML & "  if (r==true)"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
    '  sHTML = sHTML & "    alert(fr.name);"
      sHTML = sHTML & "    fr.submit();"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  else"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "    alert('Se cancelo la accion!');"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>