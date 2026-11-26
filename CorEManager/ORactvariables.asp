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
  sOPCION       = request.querystring("OPCION")
  SCODINDI       = request.querystring("INDI")






  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)










   STRSQL= "select A.*,(select c.valor from orfrecuencia c where c.objid=a.unidadfrecuen) as dias from ORindicadores a,ORUSRLST b where a.objid=b.INDICADOR and b.USRRESP='USR-"& SESSION("USUARIO") &"' AND A.ESTADO=2"


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


	SINPUT = RS(0) & "&nbsp;"
	'Celda =  ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1","","","","","si")
 	Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","10","","","","CODIGO","si")
	Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=center","","30%","","","","NOMBRE","si")
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=center","","30%","","","","DESCRIPCION","si")
     Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","10%	","","","","ESTADO","si")
    ' Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","10%","","","","TIPO","si")
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=CENTER","","10%","","","","VENCIDO","si")
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","10%","","","","ACT VARIABLES   ","si")

	Formulario = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  posI = 0





WKSETTABLA="ORindicadores"
  Do While (Not rs.EOF)
    sForm = "form" & posi

    Celda = ""
    'Tipo

     'Eliminar
	if bSuprimir and rs("tipo")=2 then
	  sHref = "ORCARGAUSRINDI.asp?OPCION=ELIM&INDI=" & rs(0)

	  Celda = ""'ofv.BotonEliminar(rs(1),sHref)
	else
	  Celda =""' ofv.GenCelda("","","","","","","","&nbsp;","si")
	end if


    SINPUT = RS(0) & "&nbsp;"

	Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    SINPUT = RS(1) & "&nbsp;"

	Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

   'descripcion


   	       nCol = 120

           wdescri=""

   	       if  not  isnull(rs(3)) then
   	           wdescri = rs(3)
   	       end if

   	       if  len(wdescri) = 0 then wdescri = " "

   	       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' "
   	       t3b = t3b & "name=" & rs(3).name & " cols=" & nCol & " rows=2 "
   	       if  bModificar then
   	           t3b = t3b & " onblur=" & sForm & ".submit() "
   	           t3b = t3b & " onkeypress=" & chr(34)
   	           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
   	           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
   	       else
   	           t3b = t3b & " readonly >"
   	       end if
   	       xdatos =  wdescri
   	       t3b    = t3b & xdatos & "</textarea></font>"
   	       SINPUT =wdescri ' t3b

           SINPUT = RS(3) & "&nbsp;"

	      Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


        SINPUT=RS(8)

        SELECT CASE SINPUT
        CASE 1
          SINPUT="NUEVO"
        CASE 2
          SINPUT="INICIADO"
        CASE 3
          SINPUT="NO VIGENTE"
        CASE ELSE
          SINPUT="SIN DEFINIR"
        END SELECT

        Celda = Celda & ofv.gencelda("","UTi  style='border-style:NONE;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


         SINPUT=RS("TIPO")

		SELECT CASE SINPUT
		CASE 1
		  SINPUT="IMPORTADO"
		CASE 2
		  SINPUT="DEL USUARIO"
		CASE ELSE
		  SINPUT="SIN DEFINIR"
		END SELECT



     'sinputVal=validarFecuencia(rs(0),rs("FECHAULTACT"))
     sinputVal=validarFecuencia2(rs(0),rs("FECHAULTACT"),rs("dias"))







  'ASIGNAR

  	sImagenP = "&#0232;"
  	sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
  	sinput = "<a class='ft1n' " & sFuncAnc & " href='ORCARGAVARIABLES.asp?ID=" & rs(0) & "&CantRegMover=" & CantRegMoverx
  	sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & "&Pendiente="&sinputVal
  	sinput = sinput  & "' title='Propiedades de " & rs(1) & "' >" & sImagenP & "</a>"






    Celda = Celda & ofv.gencelda("","UTi  style='border-style:none;border-width:1;border-color:#BBBBBB;' align=center","","","","","",sinputVal,"si")
    Celda = Celda & ofv.gencelda("","UTi  style='border-style:none;border-width:1;border-color:#BBBBBB;' align=center","","","","","",sInput,"si")




   ' sInput = ofv.gencomboFX(sForm,rs(7).name,rs(7),scomK,scomD," ")

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

  'Botonera
  Agregar = ""

  Buscar  = ""
  Celda   = ""
  Agregar = ""
  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)




  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML


function validarFecuencia2(Idindicador,fechaultact,dias)


  sinput="SI"

  FEC=cdate( dias + CLNG("0" & fechaultact) + clng(cdate("01-01-1900") ))

  if DATE()>= fec then
    SInput="SI"
   ELSE
    SINPUT="NO"
   end if
  validarFecuencia2=sinput '& fec & "--" & date()
end function


function validarFecuencia(Idindicador,fechaultact)

 STRSQL="SELECT * FROM CSORVARIABLES2	 WHERE  activa=0 and IDINDICADOR=0" & Idindicador & " order by fechact desc"
  set rspend = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    sinput="NO"
    if not rspend.eof then
    ' response.write day(DATE) & "---" &  DAY(CDATE(CLNG(rspend("fechact")) + clng(cdate("01-01-1900") )))  & chr(13)
      select case rs("unidadfrecuen")
       case 1 ' diario
           FEC=cdate( CLNG("0" & rspend("fechact")) + clng(cdate("01-01-1900") ))

	       if day(DATE()) >= day(FEC) then
	         Sinput="SI"
	       end if
	    case 2 'mensuaL
	         FEC=cdate(CLNG("0" & rspend("fechact")) + clng(cdate("01-01-1900") ))
	         if MONTH(DATE()) > MONTH(FEC) and year(date())>= year(fec) then
			   Sinput="SI"
	         end if
	    case 3 ' anual
	    end select



    else
      select case rs("unidadfrecuen")
		 case 1 ' diario
			 FEC=cdate( CLNG("0" & fechaultact) + clng(cdate("01-01-1900") ))
		   if day(DATE()) > day(FEC) then
			 Sinput="SI"
		   end if
		case 2 'mensuaL
			 FEC=cdate(CLNG("0" & fechaultact) + clng(cdate("01-01-1900") ))
			 if MONTH(DATE()) > MONTH(FEC) and year(date())>= year(fec) then
			   Sinput="SI"
			 end if
		case 3 ' anual
	  end select

    end if
    call ofv.cerrarconsulta(rspend)

    validarFecuencia=sinput

end function

%>