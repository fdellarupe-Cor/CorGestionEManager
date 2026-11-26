<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl("0" & request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  
  Formulario=""
  
  stabla=""
  
  sImageva = ""   '"<font face='corflaticon' size=2 color='#00aa00'><b>&#0252;</b></font>" & "&nbsp;&nbsp;"
      sImageIT = "<font face='corflaticon' size=2 color='#FF0000'><b>&#0167;</b></font>" & "&nbsp;&nbsp;"
  
      wkstylof = "aaa style='font-family:arial;font-size:6pt;color:#888888;' "
  
      BCOLORG = "#cccccc"
      BCOLORD = "#6699CC"
  
      '''''''''''''''''''''''''''''
  
      wkstylo1 = " style='background-color:#ddddff;color:#000000;font-weight:bold;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "
      wkstylo2 = " style='background-color:#f7eedd;color:#000000;font-weight:bold;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "
    wkstylo3 = " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' "
    
    
  idp=request.querystring("id")
 
    IF idp <> "" THEN
       SESSION("idp") = idp
    ELSE
       idp = SESSION("idp")
  END IF 
  
  
  set cn = ofv.conectar(ofv.strconn2)

  Strsql =  "Select count(*) From PUESTOREL WHERE idpst = '" & idp  & "'" 
 ' RESPONSE.WRITE Strsql
  
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    MaxRows=rs(0)
  else
    MaxRows=0
  end if
  
  call ofv.cerrarconsulta(rs)
  
  Strsql =  "Select cod From PUESTOREL WHERE idpst = '" & idp  & "' ORDER BY COD" 
    'RESPONSE.WRITE Strsql
    
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  
  if MaxRows = 1 then
  
   response.redirect "verinfoUSR.asp?id=" & rs(0) & "&IDP=" & IDP & "&bsq=USUARIO&opc=NO"
  elseif maxrows > 1 then
    Formulario = ofv.Encabezados("NOMBRE","Depto")
  elseif  maxrows=0 then
  
       SQL= "Select C.DESCRI,C2.DESCRI From  LISTADEPUESTOS C "
       SQL= SQL & "LEFT JOIN ECO C2 ON (C2.ECO=C.SECTOR) "
       SQL =SQL & "WHERE C.OBJID=" & IDP
       
       set rsF = ofv.crearconsultaEx(SQL,cn,1,parametros)

       Celda =  OFV.GenCelda("", "UTI   " & wkstylo1, "", "10%", "", "", "", "Puesto:", "si")
       sHref="GENPUESTOFUNCION.asp?PUESTO="  & IDP
       sinput=ofv.GenLink("", "cc", sHref, "", "", "", RSF(0), "si")
       
       Celda = Celda & OFV.GenCelda("", "zzz", "", "50%", "", "", "", sinput , "si")
    
        Fila =  OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")
            
         Celda = ""
    	        
    	Celda = celda & OFV.GenCelda("", "UTI   " & wkstylo1, "", "", "", "", "", "Dependencia:", "si")
    	
    	idp=traerpstpadre(idp)
    	
    	sHref="OrganigramaPuestosgraf.asp?id=" &IDP
	sinput=ofv.GenLink("", "cc", sHref, "", "", "", RSF(1), "si")
    	Celda = Celda & OFV.GenCelda("", "zzz", "", "", "", "", "", sinput, "si")
    	
        Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")
        
        Celda =  OFV.GenCelda("", "UTI   colspan=2 ALIGN=CENTER" & wkstylo1 , "", "", "", "", "", "PUESTO VACANTE", "si")
	    	
        Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")
        
        stabla="<CENTER>" & OFV.GenTabla("", "AAA style='border-style:none;border-color:#cccccc;border-width:1' ", "", "50%   ", "", "0", "", "1", "1", "", "", Fila, "si")& "</CENTER>"
        CALL OFV.CERRARCONSULTA(RSF)
  end if
  
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

   

    xx=getInfouserPst(rs(0)&"",IDP,"USUARIO")
    Celda =  ofv.gencelda("","","","","","","","","si")
    'Descripcion
    sInput = replace( session("SGU_nombre")," " ,"&nbsp;") 
    
    shref="verinfoUSR.asp?id=" &rs(0) & "&IDP=" & IDP & "&bsq=USUARIODEP"
    sinput=ofv.GenLink("", "cc", sHref, "", "", "", sInput, "si")
    'ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"130","dt","")
    Celda = celda &  ofv.gencelda("","","","60%","","","",sInput,"si")
    
    sInput=session("SGU_DEPTO")
    Celda =Celda  &  ofv.gencelda("","","","40%","","","",sInput,"si")
    
     if xTrColor = "#FFFFFF" then
                  xTrColor = "#F9F9F9"
               ELSE
                  xTrColor = "#FFFFFF"
       END IF
    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor,"","","","","",Celda,"si")
    
    
    Formulario = Formulario & FILA 'ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

 

 ' str1 = "Select * From usuarios "'where sector=" &id &" Order By descri"
  'str2 = "Select * From usuarios "'where sector=" &id &" Order By descri"
  Buscar =""' ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
 
  sAccion = "genpuestofuncion.asp?puesto=" & idp
    
  retornar = ofv.BotonVolver(sAccion)
  shtml=shtml & stabla
  IF MAXROWS > 1 THEN
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,retornar)
  else
    sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,retornar)

  END IF
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

FUNCTION  traerpstpadre(spuesto)

SQL = "select grafica from listadepuestos where objid=0" & spuesto
Set rsx2PX = OFV.crearconsultaEx(SQL, cn, 1, parametros)
if not rsx2PX.eof then
xgrafica= rsx2PX(0)
else
xgrafica="N"
end if
call ofv.cerrarconsulta(rsx2PX)

if xgrafica="S" then
        puestomaster=spuesto
     '   PUESTOMASTERdescri=spsector
    else
    SQL = "SELECT a.PUESTOMASTER,e.descri,e.grafica FROM RELAPUESTO a "
    SQL = SQL & " inner join  listadepuestos e on (e.objid=a.PUESTOMASTER) WHERE  a.PUESTODEPEN=0" & spuesto & " ORDER BY grafica "
    'response.write sql &"<br>"
    
   
    Set rsx2PX = OFV.crearconsultaEx(SQL, cn, 1, parametros)
    If Not rsx2PX.EOF  Then
        PUESTOMASTER = rsx2PX(0)
        PUESTOMASTERdescri = rsx2PX(1)
        'xgrafica=rsx2PX(2)
       ' Set rsX = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
	
	    Do While Not rsx2PX.EOF 

		PUESTOMASTER = rsx2PX("PUESTOMASTER")
		If rsx2PX("grafica") = "S" Then Exit Do
		Call OFV.cerrarconsulta(rsx2PX)
		StrSql = "Select RelaPuesto.*,listadepuestos.descri,listadepuestos.continuador,listadepuestos.GRAFICA FROM RelaPuesto  "
		StrSql = StrSql & " left join listadepuestos on  (RelaPuesto.puestomaster = listadepuestos.objid) "
		StrSql = StrSql & " WHERE RelaPuesto.puestodepen = 0" & PUESTOMASTER & "  ORDER BY RelaPuesto.puestorel "
		'response.write strsql
		Set rsx2PX = OFV.crearconsultaEx(StrSql, cn, 1, parametros)
             Loop
                    
                    
    Else
        PUESTOMASTER = spuesto
       ' PUESTOMASTERdescri = ""
    End If
   end if
  ''  If xgrafica = "N" Then
        zpuesto = PUESTOMASTER
  ''  Else
   '     zpuesto = spuesto
        'PUESTOMASTERdescri = spsector
   ' End If
         zpuesto=puestomaster
         
    traerpstpadre=zpuesto     
end function
Response.Write sHTML

%>