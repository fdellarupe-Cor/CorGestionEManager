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


  SORCODIGO   = request.FORM("ORCODIGO")
  sORTEXTO = request.FORM("ORTEXTO")
  SORESTADO    = request.FORM("ORESTADO")
  sORCLASIFICAINDI   = request.FORM("ORCLASIFICAINDI")
  SORVENCIDOS   = request.FORM("ORVENCIDOS")


  WSPSESION="AND"
  SESSION(WSPSESION) =  ""

      IF SORCODIGO <> "TODOS" THEN

         SORCODIGOW   = " (A.OBJID = 0" & SORCODIGO & ") "

         SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORCODIGOW
         SCONECTOR = " AND "

      END IF

      IF SORTEXTO <> "TODOS" THEN
         sORCARACTERW = " (A.DESCRIPCION LIKE '%" & SORTEXTO  & "%') "

         SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & sORCARACTERW
         SCONECTOR = " AND "

      END IF



      IF SORESTADO <> "TODOS" THEN
         SORTEXTOW    = " (A.ESTADO="& SORESTADO &") "

         SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORTEXTOW
         SCONECTOR = " AND "

      END IF


    IF sORCLASIFICAINDI <> "TODOS" THEN
            SORTEXTOW    = " (A.CLASIFICACION="& sORCLASIFICAINDI &") "
            SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORTEXTOW
            SCONECTOR = " AND "

         END IF







  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



  Fila = ""
  xTrColor = "#FFFFFF"

  sImageva    = "<font face='corflaticon' size=2 color='#990000'>&#0108;</font>"
  sImagevaC="&#0108;"
  sImagevm    = "<font face='corflaticon' size=2 color='#f7aa22'>&#0110;</font>"
  sImagevb    = "<font face='corflaticon' size=2 color='#009900'>&#0110;</font>"
  sImagevn    = "<font face='corflaticon' size=2 color='#dddddd'>&#0110;</font>"

  sIFORM    = "<font face='corflaticon' size=2 color='#FFFFFF' onclick=""document.all.DX5.style.display='block';"" TITLE='Ver Formula de aplicacion.' style='cursor:pointer;background-color:#000099;border-style:OUTSET;border-width:1;' >i</font>"

  wkstylof = "aaa style='font-family:arial;font-size:6pt;color:#888888;' "

''''b.codvinobj = 'VN' AND

     BCOLORG = "#AAAAAA"
     BCOLORD = "#6699CC"

  MAINPROC = ""
  MAINACTI = ""
  MAINRIES = ""

  WKTITG = "MATRIZ DE INDICADORES DE TENDENCIA DE RIESGOS"

      wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "


    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI COLSPAN=8 " & wkstylo1,"","","","","",WKTITG,"si")


    Fila = FILA & ofv.genrow("","AAA BGCOLOR='#F7EEDD' ","","","25","","",Celda,"si")



    Celda = ""
   ' Celda = Celda & ofv.gencelda("","VTI  COLSPAN=2 " & wkstylo1,"","","","","","INDICADORES","si")
   ' Celda = Celda & ofv.gencelda("","VTI   " & wkstylo1,"","","","","","CATEGORIA","si")
   '' Celda = Celda & ofv.gencelda("","VTI   " & wkstylo1,"","","","","","LINEA DE NEGOCIOS","si")
   ' Celda = Celda & ofv.gencelda("","VTI  COLSPAN=2 " & wkstylo1,"center","","","","","VARIABLES","si")
   '' Celda = Celda & ofv.gencelda("","VTI  COLSPAN=2 " & wkstylo1,"center","","","","","RESULTADO","si")
   ' Fila = FILA & ofv.genrow("","","","","15","","",Celda,"si")

  tabla1=ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:" & BCOLORG & ";' ","","95% ","","0","","1","1","","",Fila,"si")
  FILA=""

 if session(WSPSESION )<> ""  then
  swhere= " where "
 else
  swhere= "where  a.estado=2"
 end if

  STRSQL= "select * , (select count(*) from ORRELAINDIRIESGO b where b.IDINDICADOR=a.objid ) as REL FROM ORINDICADORES a  where  a.estado=2 " '& swhere & SESSION(WSPSESION)
 ' response.write strsql

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then

     posi   = 0
     coln    =0
     Celda = ""
     fila=""
     DO WHILE NOT RS.EOF
    ' fecha = GTFechadorSH(RS("FECHACREACION"))

     strsql="select * from csorvariables3 where varnombre='Z' and idindicador=0" & rs(0)
   '  response.write strsql & "<br>"
     set vars=ofv.crearconsultaEx(StrSql,cn,1,parametros)

     if not vars.eof then
      VALOR=vars("VALOR")
      FECHACT=vars("FECHACT")
     else
      VALOR=0
      FECHACT=0
     end if


   if clng("0" & VALOR)=0 and clng("0" & FECHACT)=0 then
    descripcionUM="NO EXISTEN VARIABLES"
    sImageva    = "<font face='corflaticon' size=30 color='#ff0000'>"&sImagevaC&"</font>"
   else
   '  Strsql = "Select * From orumbrales  WHERE VALOR1 >=0" & VALOR  & " AND IDINDICADOR=0" & rs(0)
     Strsql = "Select * From orumbrales  WHERE VALOR1 <=0"&VALOR& " AND VALOR2>=0"& VALOR & " AND IDINDICADOR="&rs(0) 'id
'
     '  response.write strsql & "<br>"
	    set rsum = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	    if rsum.eof then
	 	 '
	 	  sImageva    = "<font face='corflaticon' size=30 color='#ff0000'>"&sImagevaC&"</font>"
	 	  descripcionUM="FUERA DE RANGOS"
	    else
	      sImageva    = "<font face='corflaticon' size=30 color='" & rsum("color") & "'>"&sImagevaC&"</font>"
          descripcionUM=rsum("descripcion")
        end if
    end if

     info= "<u><b>CODIGO:" & RS(0)&"&nbsp;&nbsp;"& rs("descripcion") & "</b></u>"
     CELDAINFO=ofv.gencelda("","ut " ,"","","","","",info,"si")
     CELDAINFO=CELDAINFO & ofv.gencelda("","ut  rowspan=4 align=right" ,"","","","","",sImageva,"si")

     FILAinfo= ofv.genrow("","","","","","","",CELDAINFO,"si")

     info= "ULTIMA ACTUALIZACION:&nbsp;<b>" &  GTFechadorSH(FECHACT)& "</b>"

     CELDAINFO=ofv.gencelda("","ut  " ,"","","","","",info,"si")
	 'CELDAINFO=CELDAINFO & ofv.gencelda("","ut  " ,"","","","","","&nbsp;","si")

	 FILAinfo=FILAinfo & ofv.genrow("","","","","","","",CELDAINFO,"si")




   info="UMBRAL :&nbsp;<b>" & descripcionUM &"</b>&nbsp;&nbsp;&nbsp; VALOR:&nbsp;<B>"& VALOR & "</B><br>"

   CELDAINFO=ofv.gencelda("","ut  " ,"","","","","",info,"si")
   'CELDAINFO=CELDAINFO & ofv.gencelda("","ut  " ,"","","","","","&nbsp;","si")

   FILAinfo=FILAinfo & ofv.genrow("","","","","","","",CELDAINFO,"si")


   info="<B>VER RIESGOS ASOCIADOS</b>"
   sImagenP = "&#0232;"
   sFuncAnc = ""
   sImageOF   = "<font  size=2 color='#FF0000'  >" & INFO & "</font>"
   url="'ORELRIESGOS.asp?SCODINDI=" &rs(0) & "' ,'mywindow',',scrollbars=1;resizable=0;status=0,toolbar=0,width=900,height=700'"



   if CINT("0" & RS("REL"))> 0 then
    CELDAINFO =ofv.gencelda("","","LEFT onclick="&chr(34)&"window.open("&url&");"&chr(34)  & " style='cursor:pointer;'","","","","",sImageOF,"si")
   else
    CELDAINFO =ofv.gencelda("","","","","","","","&nbsp;","si")
   end if


      'CELDAINFO=ofv.gencelda("","ut  " ,"","","","","",info,"si")
     'CELDAINFO=CELDAINFO & ofv.gencelda("","ut  " ,"","","","","","&nbsp;","si")

   FILAinfo=FILAinfo & ofv.genrow("","","","","","","",CELDAINFO,"si")



   TABLAINFO=ofv.GenTabla("","aaa width=100%","","","","0","","1","1","","",FILAinfo,"si")
   call ofv.cerrarconsulta("rsum")
   call ofv.cerrarconsulta(vars)

     Celda = Celda & ofv.gencelda("","ut style='border-style:OUTSET;border-width:1;background-color:#eeffff;' " ,"","","","","",TABLAINFO,"si")

     if coln=3 then
      fila= fila &  ofv.genrow("","","","","","","",Celda,"si")
      celda=""
      coln=0
     else
      coln=coln+1
     end if
     RS.MOVENEXT

     posi=posi +1
     LOOP
  END IF

   if fila="" or celda<>"" then fila= fila & ofv.genrow("","","","","90","","",Celda,"si")

    tabla2=ofv.GenTabla("","aaa ","","95% ","","0","","1","1","","",Fila,"si")

    sHTML = sHTML & "<CENTER>" & tabla1 & tabla2 & "</center>"




  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""
  sAccion = "ORBUSCAINDI.ASP"
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>