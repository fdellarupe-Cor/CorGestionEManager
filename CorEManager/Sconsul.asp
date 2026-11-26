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
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)


'response.Write request.form & "<br><br><br>"
'response.Write request.querystring

base=request.querystring("BASE")
ing=request.querystring("ing")


  IF base <> "" THEN
       SESSION("BASE")=base
    END IF

sHTML = ofv.FormHeader(session("FormN"))

set cn = ofv.conectar(ofv.strconn2)


   Celda = Celda & ofv.GenCelda("","uti","","90%","","","","CONSULTA DE FORMULARIOS ","si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","aaa","","100% ","30","0","","-1","1","","",Fila,"si")




   Formulario = ""

    Celda = ""
    Filax  = ""

    if ing="P" then
    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","TEXTO","si")

  SINPUT = ofv.GenerarInput("buscar","","TEXT",sForm,"140","dt","")
  Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","","<BR>" & SINPUT & "<BR><BR>","si")

  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  end if
  Celda = ""
'  Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","CODIGO DEL FORMULARIO","si")

   ' SINPUT = ofv.GenerarInput("CODIGOF","0","TEXT",sForm,"40","dtN","")
  '  Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

   ' Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    Celda = ""
    sForm = "formlis"
     'CARACTER

    if ing <> "P" then
     Celda = Celda & ofv.gencelda("","tc","","","","","","ESTADO","si")

    Strsql0 = " Select * From Estados order by estado "
    Set rs0 = ofv.crearconsultaEx(Strsql0, cn, 1, parametros)
    scomK = ""'"null" & Chr(9)

    scomD = "" '"" & Chr(9)
    SCOMSEP =""
    FIRSTDAT=""
    if  not rs0.eof then
       FIRSTDAT=rs0(0)
    end if

    do while not rs0.eof
     scomK=scomK & SCOMSEP & rs0(0)
     scomd=scomd & SCOMSEP & rs0(1)
     SCOMSEP = chr(9)
     rs0.movenext
      loop

    sInput =  ofv.gencomboFX(sForm, "ESTADO", 0, scomK, scomD, "no0")

    'SINPUT = ofv.GenerarInput("ORCODIGO","TODOS","TEXT",sForm,"40","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
   end if
  Celda = ""
  Celda = Celda & ofv.gencelda("","tc","","","","","","BANCA","si")


Strsql = "Select * From FORMADMbanca "
    Set RS = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
    scomK = ""'"null" & Chr(9)

    scomD = "" '"" & Chr(9)
    SCOMSEP =""
    FIRSTDAT=""
 if  not RS.eof then
       FIRSTDAT=RS(0)
 end if

    do while not RS.eof
     scomK=scomK & SCOMSEP & RS(0)
     scomd=scomd & SCOMSEP & RS(2)
     SCOMSEP = chr(9)
     RS.movenext
      loop

    sInput =  ofv.gencomboFX(sForm, "BANCA", 0, scomK, scomD, "no0")




   '' SINPUT = ofv.GenerarInput("ORCODIGOBCRA","TODOS","TEXT",sForm,"40","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

      Celda = ""
  Celda = Celda & ofv.gencelda("","tc","","","","","","TIEMPO DE GUARDA","si")



Strsql2 = " Select * From FORMADMTGUARDA Order By OBJID "
    Set RS2 = ofv.crearconsultaEx(Strsql2, cn, 1, parametros)
    scomK = ""'"null" & Chr(9)

    scomD = "" '"" & Chr(9)
    SCOMSEP =""
    FIRSTDAT=""
 if  not RS2.eof then
       FIRSTDAT=RS2(0)
 end if

    do while not RS2.eof
     scomK=scomK & SCOMSEP & RS2(0)
     scomd=scomd & SCOMSEP & RS2(1)
     SCOMSEP = chr(9)
     RS2.movenext
      loop

    sInput =  ofv.gencomboFX(sForm, "TGUARDA", 0, scomK, scomD, "no0")



   '' SINPUT = ofv.GenerarInput("ORCODIGOBCRA","TODOS","TEXT",sForm,"40","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  Celda = ""
  Celda = Celda & ofv.gencelda("","tc","","","","","","PRODUCTO - SERVICIO","si")


Strsql3 = "Select * From TIPOPRODUCTO "
    Set RS3 = ofv.crearconsultaEx(Strsql3, cn, 1, parametros)
    scomK = ""'"null" & Chr(9)

    scomD = "" '"" & Chr(9)
    SCOMSEP =""
    FIRSTDAT=""
 if  not RS3.eof then
       FIRSTDAT=RS3(0)
 end if

    do while not RS3.eof
     scomK=scomK & SCOMSEP & RS3(0)
     scomd=scomd & SCOMSEP & RS3(1)
     SCOMSEP = chr(9)
     RS3.movenext
      loop

    sInput =  ofv.gencomboFX(sForm, "PRODU-SERVI", 0, scomK, scomD, "no0")

    'SINPUT = ofv.GenerarInput("ORCODIGOBCRA","TODOS","TEXT",sForm,"40","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


 Celda = ""
  Celda = Celda & ofv.gencelda("","tc","","","","","","PROCESO","si")


Strsql3 = "Select * From FORMADMPROCESO "
    Set RS3 = ofv.crearconsultaEx(Strsql3, cn, 1, parametros)
    scomK = ""'"null" & Chr(9)

    scomD = "" '"" & Chr(9)
    SCOMSEP =""
    FIRSTDAT=""
 if  not RS3.eof then
       FIRSTDAT=RS3(0)
 end if

    do while not RS3.eof
     scomK=scomK & SCOMSEP & RS3(0)
     scomd=scomd & SCOMSEP & RS3(2)
     SCOMSEP = chr(9)
     RS3.movenext
      loop

    sInput =  ofv.gencomboFX(sForm, "proceso", 0, scomK, scomD, "no0")

    'SINPUT = ofv.GenerarInput("ORCODIGOBCRA","TODOS","TEXT",sForm,"40","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


Celda = ""
  Celda = Celda & ofv.gencelda("","tc","","","","","","CATEGORIA DEL FORMULARIO","si")



Strsql4 = " Select * From FORMADMCATG "
    Set RS4 = ofv.crearconsultaEx(Strsql4, cn, 1, parametros)
    scomK = ""'"null" & Chr(9)

    scomD = "" '"" & Chr(9)
    SCOMSEP =""
    FIRSTDAT=""
 if  not RS4.eof then
       FIRSTDAT=RS4(0)
 end if

    do while not RS4.eof
     scomK=scomK & SCOMSEP & RS4(0)
     scomd=scomd & SCOMSEP & RS4(1)
     SCOMSEP = chr(9)
     RS4.movenext
      loop

    sInput =  ofv.gencomboFX(sForm, "CATEGORIAF", 0, scomK, scomD, "no0")

    'SINPUT = ofv.GenerarInput("ORCODIGOBCRA","TODOS","TEXT",sForm,"40","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

Celda = ""
  Celda = Celda & ofv.gencelda("","tc","","","","","","CLASE DEL FORMULARIO","si")



Strsql5 = " Select * From FORMADMTFORMU "
    Set RS5 = ofv.crearconsultaEx(Strsql5, cn, 1, parametros)
    scomK = ""'"null" & Chr(9)

    scomD = "" '"" & Chr(9)
    SCOMSEP =""
    FIRSTDAT=""
 if  not RS5.eof then
       FIRSTDAT=RS5(0)
 end if

    do while not RS5.eof
     scomK=scomK & SCOMSEP & RS5(0)
     scomd=scomd & SCOMSEP & RS5(2)
     SCOMSEP = chr(9)
     RS5.movenext
      loop

    sInput =  ofv.gencomboFX(sForm, "CLASEF", 0, scomK, scomD, "no0")

    'SINPUT = ofv.GenerarInput("ORCODIGOBCRA","TODOS","TEXT",sForm,"40","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")






Celda = ""
  Celda = Celda & ofv.gencelda("","tc","","","","","","FIRMAS DEL FORMULARIO","si")



Strsql6 = " Select * From FORMADMTFIRMA "
    Set RS6 = ofv.crearconsultaEx(Strsql6, cn, 1, parametros)
    scomK = ""   '"null" & Chr(9)

    scomD = ""   '"" & Chr(9)
    SCOMSEP =""
    FIRSTDAT=""
 if  not RS6.eof then
       FIRSTDAT=RS6(0)
 end if

    do while not RS6.eof
     scomK=scomK & SCOMSEP & RS6(0)
     scomd=scomd & SCOMSEP & RS6(1)
     SCOMSEP = chr(9)
     RS6.movenext
      loop

    sInput =  ofv.gencomboFX(sForm, "FIRMASF", 0, scomK, scomD, "no0")

    'SINPUT = ofv.GenerarInput("ORCODIGOBCRA","TODOS","TEXT",sForm,"40","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



  Celda = ""
  Celda = Celda & ofv.gencelda("","tc","","","","","","MODO DE IMPRESION","si")


      scomK = "SF" & chr(9) & "DF"
	  scomD = "SIMPLE FAZ" & chr(9) & "DOBLE FAZ"
      '' if bModificar then
        sInput = ofv.gencomboFX(sForm,"IMPRESION",0,scomK,scomD,"no0")
       'ELSE

       '' END IF
   '' SINPUT = ofv.GenerarInput("IMPRESION","SI","TEXT",sForm,"40","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



     'CARACTER

   XFNAME   = request.querystring("XFNAME")
     XFSRH    = request.querystring("XFSRH")
     XFSRD    = request.querystring("XFSRD")


     IF XFNAME = "" THEN XFNAME = "PV"
     IF XFSRD  = "" THEN XFSRD = 0
     IF XFSRH  = "" THEN XFSRH = 0

   ' PRIMERA VEZ
    IF XFNAME="PV"  THEN
       XFSRD = clng(fechaconvert(DATE())) - clng(cdate("01-01-1900"))
       XFSRH = XFSRD '+ 15

    END IF

    IF XFNAME = "FD" THEN
     XFSRD = request.form("DESDE")
      XFSRH = XFSRH
     END IF


      IF XFNAME = "FH" THEN
      XFSRD = XFSRD
      XFSRH =request.form("HASTA")
      END IF










       sImagenE = "<B><FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=4>&#0166;</FONT></B>"
       sAccion = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioD","Desde",GTCSMT,GTCFCH,GTCFCHDAT,XFSRD,4,2,0,0,0)
       sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
       sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"


       sAccionHasta = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioH","Hasta",GTCSMT,GTCFCH,GTCFCHDAT,XFSRH,4,2,0,0,0)
       sinputHasta = "&nbsp;&nbsp;<a  class=UT href=" & sAccionHasta
       sinputHasta = sinputHasta  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"



       ' CELDAS DE DESDE

      if  clng(XFSRD) = 0 then
         sInput2 = "------------"
    else
       sInput2 = cdate(clng(XFSRD) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput2)
       sInput2 = Fechador(XFSRD)
      END IF

      celda2=""

      celda2 =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","DESDE ","si")
        sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Desde",clng(cdate(XFSRD + clng(cdate("01-01-1900")))),"hidden","","20","dt","")
        csfdesde=sInput&sinputE
        celda2  = celda2 & ofv.GenCelda("","vti","","","","","",sInput&sinputE,"si")



        FILACAL=""
      FILACAL   =   ofv.GenRow("","","","","30","","",celda2,"si")


  ' CELDAS DE HASTA
      if clng(XFSRH) = 0 then
             sInput2 = "------------"
          else
             sInput2 = fechaConvert(clng(XFSRH) + clng(cdate("01-01-1900")) )
             swd = datepart("w",sInput2)
             sInput2 = Fechador(XFSRH)
         END IF



         IF XFSRD >XFSRH THEN
           msg=" onhover='validarfechas();'"
         else
           msg=""
         END IF
         celda2=""
        celda2 = celda2 & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","HASTA ","si")
        sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Hasta",fechatostr(cdate(XFSRH+ clng(cdate("01-01-1900")))),"hidden","","20","dt","")
        csfhasta= sInput&sinputHasta
        celda2  = celda2 & ofv.GenCelda("","vti","","","","","",sInput&sinputHasta,"si")

        FILACAL   = FILACAL & ofv.GenRow("","","","","30","","",celda2,"si")
celda2=""
          celda2 = celda2 & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","TIPO ","si")
 ' scomK = "Emision" & chr(9) & "Vigencia"  & chr(9) & "Publicacion"  & chr(9) & "Vencimiento" & chr(9) & "Eliminacion"
  'scomD = "Emision" & chr(9) & "Vigencia"  & chr(9) & "Publicacion"  & chr(9) & "Revision" & chr(9) & "Eliminacion"
  scomK = "Vigencia"
  scomD = "Vigencia"

  sInput = ofv.gencomboFX("subc","fecha",0,scomK,scomD,"no0")
  'sInput =  ofv.gencomboFX("", "FIRMASF", 0, scomK, scomD, "no0")
  celda2  = celda2 & ofv.GenCelda("","vti","","","","","",sInput,"si")

  'filax   = filax & ofv.GenRow("","","","","30","","",celda2,"si")




      formx    = FILACAL
 ''     sTARGET  = ""
     ''   sAccionCA    = "ListadoConsultas.asp?Menu=" & "PORFECHA" & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren



  ' ****** CODIGO PARA CALENDARIO

   sInputDesde=ofv.GenerarInput("Desde",fechaconvert(XFSRD+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
         sAccionRetornoD="SCONSUL.asp?nombre="& sNombren &"&XFNAME=FD&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha &"&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
         formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")


   sInputHasta=ofv.GenerarInput("Hasta",fechaconvert(XFSRH+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
         sAccionRetornoH="SCONSUL.asp?nombre="& sNombren &"&XFNAME=FH&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
         formRetornoH= ofv.GenForm("frmCalendarioH","",sAccionRetornoH,STARGET,"post","",sInputHasta,"si")

  ' ****** CODIGO PARA CALENDARIO


       FormularioC = ""

       FormularioC =  ofv.GenForm(sForm,"",sAccionCA,STARGET,"post","",formx,"si") &formRetornoD & formRetornoH

      FormularioC=""



'
'
'
'
'
'
'
'
    'SINPUT = ofv.GenerarInput("ORFECINID","0","TEXT",sForm,"30","dt","")
    'Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    'SINPUT = ofv.GenerarInput("ORFECINIH","0","TEXT",sForm,"30","dt","")
    'Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    'Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



     Celda = ""
          sInput=""
            IF sMODAL<> "ENCA" THEN
          '     SINPUT = "<INPUT TYPE='button' STYLE='background-color:#ff0000;color:#ffffff;' VALUE='' CLASS='BT2' " 'EXTRAER'
          '     SINPUT = SINPUT & " onclick=""document." & sform & ".action= 'respuesta.asp?modal=" & smodal & "&OPCION=VS&ESTILO=XTR';"
          '     SINPUT = SINPUT & " document." & sform & ".target='_blank'; "
          '     SINPUT = SINPUT & " document." & sform & ".submit();"" "
          '           ''  SINPUT = SINPUT & " alert(document." & sform & ".action); "
          '            '' SINPUT = SINPUT & " alert(document." & sform & ".target);"" "
          '     SINPUT = SINPUT & " >&nbsp;&nbsp;"
               Celda = Celda & ofv.gencelda("","ut  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","CENTER","","","","",SINPUT,"si")
            else
               Celda = Celda & ofv.gencelda("","ut  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","CENTER","","","","","&nbsp;","si")
            end if


            'XFSRD&XFSRH


            sInput   = "" 'ofv.GenerarInput("","Buscar","submit","","10","bt"& msg,"") 'AGREGADO PDC'
            sInput2   = ofv.GenerarInput("desde2",clng(cdate(XFSRD )),"hidden","","10","bt"& msg,"") 'AGREGADO PDC'
            sInput3   = ofv.GenerarInput("HASTA2",clng(cdate(XFSRH )),"hidden","","10","bt"& msg,"")
   ''    SINPUT = SINPUT & " alert(document." & sform & ".action); "
   ''   SINPUT = SINPUT & " alert(document." & sform & ".target);"" "
            SINPUT = SINPUT & " &nbsp;&nbsp;"

           ' Celda = Celda & ofv.gencelda("","ut colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","", sInput2  & sInput3,"si")


           ' Filax  = Filax & ofv.genrow("","AAA bgcolor='#DDDDFF' ","","","30","","",Celda,"si")
     if ing="P" THEN

        sAccion = "SEARCHex.ASP?FORMULARIOS=K&base=K&ing=P"
     ELSE

      sAccion = "RESPUESTA2.ASP?MODAL=" & SMODAL & "&CantRegMover=" & CantRegMoverx & "&MODALMAIN=" & sMODALMAIN
      sAccion = sAccion & "&FECHA=" & sFECHA & "&OPCION=SETFRM" & "&MENU=CONSGRAL" & "&Desde="&XFSRD&"&hasta="&XFSRH
    END IF


        'Â¡latabla =""

      'Formulario =  ofv.GenForm(sForm,"",sAccion,"","post","",FilaX & sinput & sInput2  & sInput3,"si") & FormularioC
       Formulario =  ofv.GenForm(sForm,"",sAccion,"","post","",FilaX & sinput & sInput2  & sInput3,"si") & FormularioC
      SINPUT = "<br> <INPUT id=bsq name=bsq TYPE='button' VALUE='BUSCAR'"  &"CLASS='BT' "
	  SINPUT = SINPUT & " onclick=""enviar('" & sform &"');"" >"


            sHTML = sHTML & "<center>" & ofv.GenTabla("","aaa","","70% ","","0","","0","0","","",Formulario ,"si") & SINPUT &  "</center>"


 sHTML = sHTML & "<script>"
 sHTML = sHTML & "function enviar(f) {"
sHTML = sHTML & " var x = document.getElementById(f);"
sHTML = sHTML & " var txt = '';"
sHTML = sHTML & " var i;var sendf ;sendf=false;"
sHTML = sHTML & " for (i = 0; i < x.length; i++) {"
 sHTML = sHTML & " if (x.elements[i].name.toUpperCase() !='HASTA2' && x.elements[i].name.toUpperCase() !='DESDE2'){"

sHTML = sHTML & "     if ( x.elements[i].value.toUpperCase() !='0')  {sendf=true; break;} ;"
sHTML = sHTML & " }} "

 sHTML = sHTML & "  if (sendf) {document.getElementById(f).submit();}else{alert('Seleccione Algún Tipo de Busqueda')}"
 sHTML = sHTML & "}"
 sHTML = sHTML & "</script>"
 'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""


   sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)


  sHTML = sHTML & GTCCALENDAR("")

  sHTML = sHTML &  "</body></html>"
end if

Response.Write sHTML




Function Fechador(wkfecha)


      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function



Function fechaConvert(wkfecha)
  Nfecha=cdate(wkfecha)
  Nfecha=dateserial(year(wkfecha),month(wkfecha),day(wkfecha))
  fechaConvert=Nfecha
end function

Function fechatostr(wkfecha)
  fechatostr=cdate(wkfecha)
  fechatostr=day(wkfecha) & "/" & month(wkfecha) & "/" & year(wkfecha)

end function


%>



<SCRIPT LANGUAGE="JavaScript">


function validarfechas(){



  alert("La fecha Desde debe ser menor a la fecha Hasta .");
  return false;

}



//-->
</SCRIPT>