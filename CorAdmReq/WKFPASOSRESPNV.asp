<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(session("FormN"))


  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 sopcion = request.querystring("opcion")
 WKF = request.querystring("WKF")

 if WKF = "" then
    WKF = session("WKF")
 else
    session("WKF") = WKF
 end if 

 if sopcion = "SI" then
    GENWKFPASO()

 end if








    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WINGDINGS COLOR=#cc0000 SIZE=3>&#0232;</FONT></B>"



  StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & WKF & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

    IDWKF = RS(0)  

    SFORM = "Usuarios"

    sInput = "CIRCUITO DE WORKFLOW"
    celdas = ofv.GenCelda("","TTT","","20%","","","",sInput,"si")


    filas = ofv.GenRow("","","","","","","",celdas,"si")
    sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")

    filas = ""

    celdas = ofv.GenCelda("","UT","","20%","","","","&nbsp;CODIGO:","si")

    sInput = RS(1)

    celdas = celdas & ofv.GenCelda("","UT","","40%","","","",sInput,"si")


    celdas = celdas & ofv.GenCelda("","UT","","20%","","","","&nbsp;CLASE:","si")

    select case rs(5)

           case "E"

                 sInput = "EFECTIVO"

           case "M"

                 sInput = "MODELO"

           case else 
           
                 sInput = "sin definir"


    end select 


    celdas = celdas & ofv.GenCelda("","UT","","40%","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","UT","","","","","","&nbsp;DESCRIPCION:","si")

    sInput = RS(2)

    celdas = celdas & ofv.GenCelda("","UT","","","","","",sInput,"si")


    celdas = celdas & ofv.GenCelda("","UT","","","","","","&nbsp;TIPO:","si")


    strsql = "Select descripcion From WKFTIPOS where tipo = '" & rs(6) & "' "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
       sInput = rs2(0)

    else
        sInput = rs(6)
    end if


    celdas = celdas & ofv.GenCelda("","UT","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    celdas = ofv.GenCelda("","UT","","","","","","&nbsp;INICIO:","si")


    select case rs(4)

           case "A"

                 sInput = "AUTOMATICO"

           case "S"

                 sInput = "A SOLICITUD"

           case else 
           
                 sInput = "sin definir"


    end select 

    celdas = celdas & ofv.GenCelda("","UT","","","","","",sInput,"si")





    celdas = celdas & ofv.GenCelda("","UT","","","","","","&nbsp;NIVELES:","si")


    WKFRESP = "S"

    scomK = "2" & chr(9) & "3" & chr(9) & "4" & chr(9) & "5" & chr(9) & "6" & chr(9) & "7" & chr(9) & "8" & chr(9) & "9" & chr(9) & "10" & chr(9) & "11" & chr(9) & "12" & chr(9)
    scomK = scomK  & "13" & chr(9) & "14" & chr(9) & "15" & chr(9) & "16" & chr(9) & "17" & chr(9) & "18" & chr(9) & "19" & chr(9) & "20" & chr(9) & "21" & chr(9) & "22" & chr(9) & "23" & chr(9) & "24"


    scomD = scomK

    IF RS(10) = 0 THEN
       scomK = "0" & chr(9) & scomK
       scomD = "SIN ESTABLECER" & chr(9) & scomD
       WKFRESP = "N"
    end if

    WKFRESPcant = RS(10)

    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(10).name,rs(10),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(10),"readonly","","70","dt","")
    end if


    celdas = celdas & ofv.GenCelda("","UT","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(IDWKF," ","*") & "&volver=" & vuelta & "&Tabla=WKFCIRCUITO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    filas = ofv.GenRow("","","","","","","",formulario0,"si")

  sHTML = sHTML  & ofv.GenTabla("","","","100% STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","","0","","1","-1","","",filas,"si") & "<BR>"


  END IF



    celdas = ""



  'Encabezados
  formulario =  ofv.GenCelda("","ttR colspan=7","center","","","","","NIVELES DE RESPONSABLES","si")
  formulario = ofv.GenRow("","","","","22","","",formulario,"si")


  IF WKFRESP <> "S" THEN


     CELDA =  ofv.GenCelda("","VT colspan=7  ","center","","","","","&nbsp;","si")
     formulario = formulario & ofv.GenRow("","","","","60","","",CELDA,"si")



     CELDA =  ofv.GenCelda("","VT colspan=7 STYLE='border-style:solid;border-width:2;border-color:#993333;' ","center","","","","","DEBE ESTABLECER LA CANTIDAD <BR> DE NIVELES PARA EL CIRCUITO","si")
     formulario = formulario & ofv.GenRow("","","","","60","","",CELDA,"si")

  ELSE




  StrSql = "select count(NIVELRSP) from WKFPSCRESPNV WHERE WKFCOD = '" & WKF & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF not RS.EOF THEN

     
     if  WKFRESPcant < RS(0) then
''''''''''' borro niveles

         StrSql = "delete from WKFPSCRESPNV WHERE WKFCOD = '" & WKF & "' and NIVELRSP > " & WKFRESPcant
         call ofv.crearconsultaEx(StrSql,cn,1,parametros)


     elseif   WKFRESPcant > RS(0) then
''''''''''' agrego niveles


            WKDESDE = RS(0)

         for wkind=1 to (WKFRESPcant - RS(0))

             WKDESDE = WKDESDE + 1

             WKFCODPRTY = GTWKFPSC() & "N" & WKDESDE
             WKRSPDESC  = "RESPONSABLE NIVEL " &  WKDESDE
             WKRSPTIPO  = ""

             IF WKDESDE = 1 THEN

                WKRSPDESC  = "SOLICITANTE"
                WKRSPTIPO  = "SOL"

             ELSEIF WKDESDE = 2 THEN

                WKRSPDESC  = "DESTINATARIO/RESPONSABLE"
                WKRSPTIPO  = "DES"

             END IF 



             strsql = "insert into WKFPSCRESPNV ("
             strsql = strsql & "WKFCOD,WKFPASRSPCOD,DESCRIPCION,NIVELRSP,"
             strsql = strsql & "MODORSP,ESTILORSP,SUPNIVELRSP,AUTOESCALARSP,"
             strsql = strsql & "TIPORSP,CODIGORSP"
             strsql = strsql & ") values ("
             strsql = strsql & "'" & session("WKF") & "','" & WKFCODPRTY & "','" & WKRSPDESC & "','" & WKDESDE & "',"
             strsql = strsql & "'U','A','8','S',"
             strsql = strsql & "'" & WKRSPTIPO & "',''"
             strsql = strsql & ")"
 '   response.write strsql & "<br>" & chr(13)
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)


         next 



     end if




     StrSql = "select * from WKFPSCRESPNV WHERE WKFCOD = '" & WKF & "' ORDER BY NIVELRSP"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  END IF

    xTrColor = "#FFFFFF"


  posi = 0
  Do While (Not rs.EOF) 

     sform = "WKF" & rs(0)

     sc = "UT"

     sinput = RS(4) & ""
     celdas = ofv.GenCelda("","vti  bgcolor=#F7AA22 ","","","","","",sinput,"si")


     if bModificar then
        sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sform,"60","dt","")
     else
        sInput = ofv.GenerarInput("",rs(3),"readonly","","60","dt","")
     end if
     celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")



        sinputK = ""
        sinputt = ""

    scomK = "U" & chr(9) & "L" & chr(9) & "LD"
    scomD = "UNITARIO" & chr(9) & "LISTA" & chr(9) & "LISTA Y USUARIOS DEFAULT"

    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(5).name,rs(5),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(5),"readonly","","70","dt","")
    end if

    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")



        sinputK = ""
        sinputt = ""

    scomK = "F" & chr(9) & "A" & chr(9) & "M"
    scomD = "PREFIJADO" & chr(9) & "ASIGNAR POR APL ORIGEN" & chr(9) & "MANUAL ASIGNADO POR EL RESPONSABLE"

    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(6).name,rs(6),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(6),"readonly","","70","dt","")
    end if
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


    sc = "Ut  bgcolor=#EEEEEE "



    sinput = RS(7)
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    sinput = RS(8)
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    sinput = RS(9)
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    sinput = RS(10)
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    filas = ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","25","","",celdas,"si")


       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0" 
       ELSE
          xTrColor = "#FFFFFF" 
       END IF



    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPSCRESPNV"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if


    posi = posi + 1
    rs.Movenext
  loop


  
  END IF

  sHTML = sHTML  & ofv.GenTabla("","","","100%","","0","","-2","-1","","",formulario,"si")




  'Botonera
  Agregar = ""
  Buscar = ""
  
  sAccion = "WKFPASOSNV.asp?WKF=" & wkf
  Celda = ofv.BotonVolver(sAccion)


 sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Celda,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML


%>