<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  Server.ScriptTimeout=360
  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sOPCION        = request.querystring("OPCION")
  stevento        = request.querystring("TEVENTO")

  if sMODAL = "" then sMODAL = "CSN"

  sHTML = ofv.MENUHeader("","")

  WKVISION = ""



  WSPSESION = "ORWHERE"




  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  DIM TABDISTRIB(100,5)


  DIM TABRIESGOS(110,110)

  TTABRIESGOS1 = 0
  TTABRIESGOS2 = 0

  ITABRIESGOS1 = 0
  ITABRIESGOS2 = 0

          TABRIESGOS(10,0) = 0             'CODIGO
          TABRIESGOS(10,1) = "SIN ASIGNAR" 'DESCRI
          TABRIESGOS(10,2) = 0             'CANTIDAD
          TABRIESGOS(10,4) = 0             'total

          TABRIESGOS(0,10) = 0             'CODIGO
          TABRIESGOS(1,10) = "SIN ASIGNAR" 'DESCRI
          TABRIESGOS(2,10) = 0             'CANTIDAD
          TABRIESGOS(4,10) = 0             'total

  Strsql = "Select * From ORTIPOEVENTO ORDER BY SECUEN, DESCRI"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 10
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRIESGOS(POSI,0) = RS(0) 'CODIGO
          TABRIESGOS(POSI,1) = RS(1) 'DESCRI
          TABRIESGOS(POSI,2) = 0     'CANTIDAD
          TABRIESGOS(POSI,4) = 0     'total


          TTABRIESGOS1 = POSI

        RS.MOVENEXT
     LOOP



  END IF



  Strsql = "Select * From ORLINNEGO1  ORDER BY DESCRIpcion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 10
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRIESGOS(0,POSI) = RS(0) 'CODIGO
          TABRIESGOS(1,POSI) = RS(1) 'DESCRI
          TABRIESGOS(2,POSI) = 0     'CANTIDAD
          TABRIESGOS(4,POSI) = 0     'total


          TTABRIESGOS2 = POSI


          for ii=10 to  TTABRIESGOS1
              TABRIESGOS(ii,POSI) = 0     'monto
          next

        RS.MOVENEXT
     LOOP



  END IF

  TABRIESGOS(10,10) = 0     'monto






    WSPWHERE = SESSION(WSPSESION) & " AND (A.TRCEVENTORIESGO1 = 0" & sTEVENTO & ") "








  if sMODAL = "CSN"  then




     Strsql = "     select A.*, B.ACTIVIDAD FROM ORTRACER A "
     Strsql = Strsql & " LEFT JOIN (SELECT DISTINCT X.ACTIVIDAD FROM ORACTIVRIESGO X)  B ON (A.trcidentact = B.ACTIVIDAD)  "
     Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCESTADO desc , A.FECREGIS DESC  "
     WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"

     if sMODAL = "CSN"  then      WKACCION = " - LISTA DE EVENTOS RAIZ"


	'response.write chr(13) & chr(13)
	'response.write strsql
	'response.write chr(13) & chr(13)





    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg COLSPAN=9 BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",session("FormN") & " - DETALLE DE EVENTOS","si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")



    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","center","7%","","","","CODIGO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","center","20%","","","","DESCRIPCION","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","8%","","","","MONTO BRUTO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","8%","","","","OTROS MONTOS","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","8%","","","","PERDIDAS","si")
	Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","8%","","","","NO CONTABLE","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","8%","","","","RECUPEROS","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","8%","","","","DESAFECT","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","8%","","","","MONTO NETO","si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")


            WSPIMPOFTT = 0
            WSPCANTFTT = 0

        wkmontonetoT =  0
        wkmontoRECUT =  0
        wkmontoBRUTT =  0
        wkmontoOMONT =  0
        wkmontoPERDT =  0
        wkmontoDESAT =  0
		wkmontoNOCONT =  0




  'response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


  ELSE
  posI = 0

  xTrColor = "#f0f0f0"



  posi = 0
  Do While (Not rs.EOF)

	'response.Write(chr(13))
	'response.Write(rs(42))
	'response.Write(chr(13))

     wkmontonetox =  rs(16) & ""

     if wkmontonetox =  "" OR wkmontonetox =  "0"  then


     else


        xTrColor = "#f0f0f0"



        wkmontoneto =  cdbl(rs(16))


        wkmontoRECU =  cdbl(rs(15))
        wkmontoBRUT =  cdbl(rs(12))
        wkmontoOMON =  cdbl(rs(14))
        wkmontoPERD =  cdbl(rs(13))
        wkmontoDESA =  cdbl("0" & rs(36))
		'wkmontoNOCON =  cdbl("0" & rs(42))
		if rs(42)&""="" then
		  wkmontoNOCON =0
		  else

	       wkmontoNOCON =  cdbl(rs(42))
		end if



        wkmontonetoT =  wkmontonetoT + wkmontoneto
        wkmontoRECUT =  wkmontoRECUT + wkmontoRECU
        wkmontoBRUTT =  wkmontoBRUTT + wkmontoBRUT
        wkmontoOMONT =  wkmontoOMONT + wkmontoOMON
        wkmontoPERDT =  wkmontoPERDT + wkmontoPERD
        wkmontoDESAT =  wkmontoDESAT + wkmontoDESA
		wkmontoNOCONT = wkmontoNOCONT + wkmontoNOCON



        CELDA1 = ""


        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","",RS(0),"si")


        WKMDESCRI = RS(11) & "."

        Celda1 = Celda1 & ofv.gencelda("","UTI style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","","","","",WKMDESCRI,"si")

            wspcolor = "#eeeeff"

            WSPIMPOF = wkmontoBRUT


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")


            WSPIMPOF = wkmontoOMON


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

            WSPIMPOF = wkmontoPERD


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

	'***************************************************************************************
	'*
	'***************************************************************************************

            WSPIMPOF = wkmontoNOCON


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")



	'***************************************************************************************
	'*
	'***************************************************************************************

            WSPIMPOF = wkmontoRECU


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

            WSPIMPOF = wkmontoDESA


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

            WSPIMPOF = wkmontoneto


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")




        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#DDDDDD' ","","","25","","",Celda1,"si")



''''''''' BUSCA VINCULADOS '''''''''''''''''''''


     Strsql = "     select A.*, B.ACTIVIDAD FROM ORTRACER A "
     Strsql = Strsql & " LEFT JOIN (SELECT DISTINCT X.ACTIVIDAD FROM ORACTIVRIESGO X)  B ON (A.trcidentact = B.ACTIVIDAD)  "
     Strsql = Strsql & " WHERE A.TRCCODTRACERP = '" & RS("CODTRACER") & "'  ORDER BY A.TRCESTADO desc , A.FECREGIS DESC  "
     'RESPONSE.WRITE STRSQL

	 'response.Write(chr(13) & chr(13))
	 'response.Write(STRSQL)
	 'response.Write(chr(13) & chr(13))
     set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rsX1.eof then


     ELSE
     posi = 0
     Do While (Not rsX1.EOF)

        'wkmontoneto =  cdbl("0" & rsX1(16))
		wkmontoneto =  0


        wkmontoRECU =  cdbl(rsX1(15))
        wkmontoBRUT =  cdbl(rsX1(12))
        wkmontoOMON =  cdbl(rsX1(14))
        wkmontoPERD =  cdbl(rsX1(13))
        wkmontoDESA =  cdbl("0" & rsX1(36))
		'wkmontoNOCON = cdbl("0" & rsX1(42))
		wkmontoNOCON = 0







        CELDA1 = ""


        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#F7AA22' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","",RSX1(0),"si")


        WKMDESCRI = RSX1(11) & "."

        Celda1 = Celda1 & ofv.gencelda("","UTI style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","","","","",WKMDESCRI,"si")

            wspcolor = "#eeeeff"

            WSPIMPOF = wkmontoBRUT


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")


            WSPIMPOF = wkmontoOMON


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

            WSPIMPOF = wkmontoPERD


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")



	'***************************************************************************************
	'*
	'***************************************************************************************

            WSPIMPOF = wkmontoNOCON


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")



	'***************************************************************************************
	'*
	'***************************************************************************************

            WSPIMPOF = wkmontoRECU


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

            WSPIMPOF = wkmontoDESA


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

            WSPIMPOF = wkmontoneto


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")




        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#FFFFFF' ","","","25","","",Celda1,"si")


     rsX1.Movenext
     Loop

'''''''''''''''''''''''''''''


     end if

     end if



    posi = posi + 1
    rs.Movenext
  Loop




        CELDA1 = ""



        WKMDESCRI = "TOTAL:"

        Celda1 = Celda1 & ofv.gencelda("","VTI COLSPAN=2 BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","",WKMDESCRI,"si")

            wspcolor = "#eeeeff"

            WSPIMPOF = wkmontoBRUTT


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")


            WSPIMPOF = wkmontoOMONT


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

            WSPIMPOF = wkmontoPERDT


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

	'***************************************************************************************
	'*
	'***************************************************************************************

            WSPIMPOF = wkmontoNOCONT


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")



	'***************************************************************************************
	'*
	'***************************************************************************************



            WSPIMPOF = wkmontoRECUT


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

            WSPIMPOF = wkmontoDESAT


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

            WSPIMPOF = wkmontonetoT


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#EEF7DD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")




        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")







    sHTML = sHTML & "<BR>"

    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","80% ","","0","","1","1","","",Fila1,"si") & "</CENTER>"


  END IF





  else



  end if




'''bgcolor='#ddddff'
''#F7CC88

       wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "



  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sAccion = "ORMAPAEVENTOS.asp?modal=" & smodal & "&CODTRACER=" & SCODTRACER
  Celda = "" 'ofv.BotonVolver(sAccion)


  sHTML = sHTML &  ofv.botoneraESSX("0","",cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)


  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if


function TablaBusqueda()

TablaBusqueda = ""

end function
Response.Write sHTML



%>