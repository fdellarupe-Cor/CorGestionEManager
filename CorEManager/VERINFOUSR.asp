
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok = OFV.CheckUsuario()
If ok <> "Y" Then
    uv = OFV.uv
    If uv = "NO" Then ok = OFV.ValidarUsuario
    shtml = ok
Else
    response.expires = 0
    'response.write session("codSol")
    OFV.ObtenerAtributos session("Form2"), ""
    bModificar = OFV.AModificar
    If session("Nomodificar") = "S" Then bModificar = False


    sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

    ID = request.querystring("ID")
    IDP = request.querystring("IDP")
    opc= request.querystring("opc")
    bsq=  request.querystring("bsq")
    CantRegMoverx=request.querystring("CantRegMover")

  ' response.write request.querystring
    xTrColor = "#FFFFFF"

    sImageva = ""   '"<font face='corflaticon' size=2 color='#00aa00'><b>&#0252;</b></font>" & "&nbsp;&nbsp;"
    sImageIT = "<font face='corflaticon' size=2 color='#FF0000'><b>&#0167;</b></font>" & "&nbsp;&nbsp;"

    wkstylof = "aaa style='font-family:arial;font-size:6pt;color:#888888;' "

    BCOLORG = "#cccccc"
    BCOLORD = "#6699CC"

    '''''''''''''''''''''''''''''

    wkstylo1 = " style='background-color:#ddddff;color:#000000;font-weight:bold;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "
    wkstylo2 = " style='background-color:#f7eedd;color:#000000;font-weight:bold;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "
    wkstylo3 = " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' "


   if bsQ <> "USUARIO" then
        SACCION="GENPUESTOFUNCION.ASP?PUESTO="&IDP
        BTIT1 = "<input  type=BUTTON   value='Funciones del Puesto' class=bietxww  STYLE='width:100%;cursor:pointer;color:"&session("CMFMEFT")&";background-color:"& Session("CMFMENU") &"' onclick='document.getElementById('TestFrm1').submit();'   >"
        BTIT1 = BTIT1 & OFV.genform("TestFrm1", "AAA STYLE='display:none;' ", sAccion, "", "post target=_self", "", "", "si")


        SACCION="GENPUESTOFIRMAS.ASP?puesto="&idp
        BTIT2 = "<input  type=BUTTON    value='Régimen de Firmas Asociado' class=bietxww  STYLE='width:100%;cursor:pointer;color:"&session("CMFMEFT")&";background-color:"& Session("CMFMENU") &"' onclick='document.getElementById('TestFrm2').submit();'   >"
        BTIT2 = BTIT2 & OFV.genform("TestFrm2", "AAA STYLE='display:none;' ", sAccion, "", "post target=_self", "", "", "si")



        Celda = OFV.GenCelda("", "bietxdq   " & swkstylo4, "", "", "", "", "", BTIT1, "si")
        Celda = celda & OFV.GenCelda("", "bietxdw   " & awkstylo4, "", "10%", "", "", "", "&nbsp;", "si")

        Celda = CELDA & OFV.GenCelda("", "bietxdw   " & wwkstylo4, "", "", "", "", "", BTIT2, "si")
        Celda = celda & OFV.GenCelda("", "bietxdw   " & awkstylo4, "", "7%", "", "", "", "&nbsp;", "si")
        Fila = OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")

       shtml = shtml & "<center><P ALIGN=RIGHT> " & OFV.GenTabla("", "AAA    xstyle='border-style:solid;border-color:#cccccc;border-width:1;background-color:#888888;' ", "", "40%", "", "0", "", "1", "1", "", "", Fila, "si") &"</P></center>"
   end if


    Set cn = OFV.conectar(OFV.strconn2)


    'response.write "<center>"



    Fila = ""


              XXXX = GETFECH()


              XLFECDAT2  = XLFECDAT
              XLHORDAT2  = XLHORDAT


   ' strsql ="select * from usuarios where id='" & id & "'"
   ' response.write strsql
'    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, parametros)




        DESCRIUSR=""
        if bsq="USUARIO" THEN
         XSQL="    Select idpst,DESCRIUSR From PUESTOREL where  cod='" &  UCASE(ID & "") & "'"
    '     response.write xsql & "<br>"
	   Set xrs2C = OFV.crearconsultaEx(XSQL, cn, 1, "")
           if not xrs2C.EOF THEN
           IDP=xrs2C(0)
           DESCRIUSR=xrs2C(1)
          ELSE
            IDP=-1
          END IF
          CALL OFV.CERRARCONSULTA(XRS2C)

          xsql=" Select Ecoperfil.*,listadepuestos.descri From Ecoperfil, listadepuestos where Ecoperfil.codsector = listadepuestos.objid and codsector= '"& IDP &"' and Ecoperfil.codperfil = '****PUESTO' "
	     'response.write xsql
	      Set xrs2C = OFV.crearconsultaEx(XSQL, cn, 1, "")
	      zperfil=0
	      if not xrs2C.eof then
	       zperfil=xrs2C(0)
	      end if
	      CALL OFV.CERRARCONSULTA(XRS2C)

	      xsql= " select objid from ECOPERFILDEP where ecoperfil=" & zperfil
	      xsql= xsql & " union all "

	       xsql= xsql & " select objid from ECOPERFILDSC where ecoperfil=" & zperfil
	      xsql= xsql & " union all "

	       xsql= xsql &  " select objid from ECOPERFILFUN where ecoperfil=" & zperfil
	       xsql= xsql & " union all "

	       xsql= xsql & " select objid from ECOPERFILCOM where ecoperfil=" & zperfil

	      xsql= xsql & " union all "

	      xsql= xsql & " select objid from ECOPERFILFIR where ecoperfil=" & zperfil

	      xsql= xsql & " union all "
	      XSQL =xsql &  " select puestorel  from relapuesto where puestodepen="& idp &" or puestomaster=" & idp
	     ' response.write xsql
	      	Set xrs2C2 = OFV.crearconsultaEx(XSQL, cn, 1, "")

	      	if xrs2C2.eof then
	       	  muestra="NO"
	      	else
	      	  MUESTRA="SI"
	      	end if
	 CALL OFV.CERRARCONSULTA(xrs2C2)

       '    RESPONSE.WRITE  XSQL
        END IF

        if bsq="USUARIODEP" then
          MUESTRA="SI"
          XSQL="    Select idpst,DESCRIUSR From PUESTOREL where  cod='" &  UCASE(ID & "") & "'"
          Set xrs2C = OFV.crearconsultaEx(XSQL, cn, 1, "")
	  if not xrs2C.EOF THEN

           DESCRIUSR=xrs2C(1)
         end if
          CALL OFV.CERRARCONSULTA(XRS2C)
        end if
        XX=getInfouserPst(ID,IDP,"USUARIO")

        set fs=Server.CreateObject("Scripting.FileSystemObject")
        'response.write "EMPJPG/" & session("SGU_IMG")
         FileName = session ("DirUpload") & "Estructura\FOTOS EMPLEADOS\" & session("SGU_IMG") 'Server.MapPath(session ("DirUpload") & "Estructura\FOTOS EMPLEADOS" & session("SGU_IMG"))
       '  response.write fs.FileExists(FileName) 'FileName
       if fs.FileExists(FileName)=true then

        SIMG="<DIV><IMG ID=IT  SRC='Estructura/FOTOS EMPLEADOS/"& session("SGU_IMG")& "'></DIV>"
       else

        SIMG="<IMG  SRC='Estructura/FOTOS EMPLEADOS/sinfoto.jpg'>"
       end if
        set fs=nothing

        Celda = OFV.GenCelda("", "zzzz rowspan=8  ID=TD1 ", "", "1%", "", "", "", simg, "si")

        Celda = celda & OFV.GenCelda("", "bietx   " & qwkstylo1, "", "10%", "", "", "", "Nombre:", "si")

        Celda = Celda & OFV.GenCelda("", "zzz", "", "50%", "", "", "", session("SGU_nombre") , "si")

        Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")

        Celda = ""

	Celda = celda & OFV.GenCelda("", "bietx   " & swkstylo1, "", "", "", "", "", "Puesto:", "si")



	sHref="GENPUESTOFUNCION.asp?PUESTO="  & session("SGU_codigopuesto")

	if muestra="NO" AND session("SGU_descripuesto") <> "NO"  THEN
	 sinput="<font face=verdana color=#000080 size=1><b>" &session("SGU_descripuesto") &"<b></font>"
	ELSEIF session("SGU_descripuesto") <> "NO" AND MUESTRA="SI" THEN
	sinput=ofv.GenLink("", "cc", sHref, "", "", "", "<font face=verdana color=#000080 size=1><b><U>" &session("SGU_descripuesto") &"</u><b></font>", "si")
	ELSE
	SINPUT="Sin dato de Estructura" 'session("SGU_descripuesto")
	END IF

	Celda = Celda & OFV.GenCelda("", "zzz", "", "", "", "", "", sinput, "si")

        Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")

     '   response.write muestra & "  " & IDP
        Celda = ""
     if muestra <> "NO" or IDP=-1 then
	Celda = celda & OFV.GenCelda("", "bietx   " & xwkstylo1, "", "", "", "", "", "Dependencia:", "si")

	 xpst=traerpstpadre(session("SGU_codigopuesto"))
	sHref="OrganigramaPuestosgraf.asp?id=" & xpst 'session("SGU_codigopuesto")

	IF session("SGU_descrisector") <> "NO" AND MUESTRA="SI" THEN
	sinput=ofv.GenLink("", "nt", sHref, "", "", "", "<font face=verdana color=#000080 size=1><b><U>" &session("SGU_descrisector") & "</u><b></font>", "si")
	ELSE
	SINPUT="Sin dato de Estructura "'session("SGU_descrisector")
	END IF

	Celda = Celda & OFV.GenCelda("", "zzz", "", "", "", "", "", sinput, "si")

        Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")
     end if
        Celda = ""

	Celda = celda & OFV.GenCelda("", "bietx   " & xwkstylo1, "", "", "", "", "", "Dpto.:", "si")

	Celda = Celda & OFV.GenCelda("", "zzz", "", "", "", "", "", session("SGU_DEPTO"), "si")

        Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")

       Celda = ""

	Celda = celda & OFV.GenCelda("", "bietx   " & xwkstylo1, "", "", "", "", "", "Dirección:", "si")

	Celda = Celda & OFV.GenCelda("", "zzz", "", "", "", "", "", session("SGU_direccion"), "si")

        Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")

        Celda = ""

	Celda = celda & OFV.GenCelda("", "bietx   " & xwkstylo1, "", "", "", "", "", "Teléfono:", "si")

	Celda = Celda & OFV.GenCelda("", "zzz", "", "", "", "", "", session("SGU_telefono"), "si")

        Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")


        Celda = ""

	Celda = celda & OFV.GenCelda("", "bietx   " & zwkstylo1, "", "", "", "", "", "Correo electrónico:", "si")

	Celda = Celda & OFV.GenCelda("", "zzz", "", "", "", "", "", "<A CLASS=CC HREF=mailto:" & session("SGU_mail")& "><font face=verdana color=#000080 size=1><b><U>"&session("SGU_mail") & "</U></A>", "si")

	Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")

         Celda = ""
	if  trim (DESCRIUSR &"")<> ""	then
	Celda = celda & OFV.GenCelda("", "bietx   " & xwkstylo1, "", "", "", "", "", "Conoceme:", "si")


	Celda = Celda & OFV.GenCelda("", "zzz", "", "", "", "", "", REPLACE(DESCRIUSR&"",CHR(13),"<BR>"), "si")

	Fila = Fila & OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")
        end if






   ' End If


    shtml = shtml &  OFV.GenTabla("", "AAA    style='border-style:none;border-color:#cccccc;border-width:1' ", "", "95%   ", "", "0", "", "1", "1", "", "", Fila, "si")


     '' BTIT1 = "<A CLASS=CC HREF='BSQINFO.ASP?BSQ=PORPUESTO&ID="&spuesto&"'>Ver Datos de los Ocupantes del Puesto</A>"
     ' BTIT2 = "<A CLASS=CC HREF='BSQINFO.ASP?BSQ='>Ver el Régimen de Firmas Asociado al Puesto</A>"

     'Celda = OFV.GenCelda("", "UTI   " & wkstylo1, "", "", "", "", "", BTIT1, "si")
     'Celda = CELDA & OFV.GenCelda("", "UTI   " & wkstylo1, "", "", "", "", "", BTIT2, "si")

   '  Fila = OFV.genrow("", "AAA  ", "", "", "25", "", "", Celda, "si")

   ' shtml = shtml & "<BR><center>" & OFV.GenTabla("", "AAA style='background-color:#888888;' ", "", "", "", "0", "", "1", "1", "", "", Fila, "si") & "</center>"



             XXXX = GETOBJALOGUPD("ACCMOD")

     Agregar = ""
     Buscar  = ""
     Celda   = ""
   'response.write request.querystring
     if opc <> "NO" then
      if bsQ="USUARIO" then
      XGUFILTRO=request.querystring("XGUFILTRO")
      XGUITEMCD=request.querystring("XGUITEMCD")
      sAccion ="bsqOrganigrama.asp?bsq=PORNOMBRE&XGUFILTRO=" &XGUFILTRO & "&XGUITEMCD=" & XGUITEMCD

      else
      sAccion = "bsqinfo.asp?BSQ=USUARIO&ID=" &IDP &"&CantRegMover=" & CantRegMoverx

      end if
   ' if bsQ <> "USUARIODEP" then
    Celda = ofv.BotonVolver(sAccion)
  '  end if
    end if

   sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",0,"","","no",Agregar,Buscar,celda)




    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)

  '  SHTML=SHTML & "<script>"
   ' SHTML=SHTML & "var newImage = document.createElement('img');"
   ' SHTML=SHTML & "newImage.src='EMPJPG/"& id &".jpg';newImage.width='10';"
   ' SHTML=SHTML & "var elem2 = document.getElementById('TD1');"
   ' SHTML=SHTML & "elem2.appendChild(newImage);"

   ' SHTML=SHTML & "</script>"
    shtml = shtml & "</BODY>"

    shtml = shtml & "</HTML>"
End If

FUNCTION  traerpstpadre(spuesto)

SQL = "select grafica from listadepuestos where objid=0" & spuesto
'response.write request.querystring
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
response.write shtml



%>
