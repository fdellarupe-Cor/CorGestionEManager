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

  ACampos=array("FECHA&nbsp;CARGA#FECREGIS","FECHA&nbsp;DESCUBRIMIENTO#FECIDENT","FECHA&nbsp;INICIO#FECORIG","FECHA&nbsp;CONTABLE#FECCONTA","FECHA&nbsp;ESTADO&nbsp;FINALIZADO#FECFINAL")

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
 ''''' Formulario = ofv.Encabezados("CODIGO&nbsp;DISEÑO","CODIGO&nbsp;ENTIDAD")

  Formulario = ""

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



  strsql= "SELECT * FROM ORCONFIGBCRA"

  SET RS=OFV.crearconsultaEx(StrSql,cn,1,volver)

  IF RS.EOF THEN
  'IF SOPCION = "SI" THEN

        strsql = "Insert into ORCONFIGBCRA "
        strsql = strsql & " (CODDISENO,FECREGIS,FECIDENT,FECORIG,FECCONTA,"
        strsql = strsql & " FECFINAL,TOPEMONTO,ENCAMENORES,REPESUMA,GENERAIDPRES,CODIGOCUP)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('8501','S','N','N','N','N','1000','S','S','S','CUP')"

  'response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF
  CALL OFV.CERRARCONSULTA(RS)

  WKSETTABLA = "ORCONFIGBCRA"

  Strsql = "Select * From " & WKSETTABLA
  'RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    scomK = "S" & chr(9) & "N"
    scomD = "SI" & chr(9) & "NO"

  posI = 0
  Do While (Not rs.EOF)
    sForm = "form" & posi

    Fila = ""

    CELDA=ofv.gencelda("","TT COLSPAN=3","","","","","","<u>CONFIGURACION DE LA GENERACION</u>","si")
    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")

    CELDA=ofv.gencelda("","VT COLSPAN=3","","","","","","&nbsp;","si")
    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")



        celda=ofv.GenCelda("","","","","","","","&nbsp;","si") & ofv.GenCelda("","VT","","70%","","","","CODIGO DISEÑO:","si")
        sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"10","dt","")

        Celda = Celda & ofv.gencelda("","UT","","30%","","","",sInput,"si")
        Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


        celda=ofv.GenCelda("","","","","","","","&nbsp;","si") & ofv.GenCelda("","VT","","","","","","CODIGO ENTIDAD:","si")
        sInput = ofv.GenerarInput(rs(2).name & " maxlength=5" ,rs(2) ,"text",sForm,"30 ","dt","")

        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")
        Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

        celda=ofv.GenCelda("","","","","","","","&nbsp;","si") & ofv.GenCelda("","VT","","","","","","GENERA ENCABEZADO Y ID DE PRESENTACION:","si")
        sInput = ofv.gencomboFX(sForm,RS(11).NAME,rs(11),scomK,scomD," ")

        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")
        Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


        celda=ofv.GenCelda("","","","","","","","&nbsp;","si") & ofv.GenCelda("","VT","","","","","","CODIGO CUP:","si")
        sInput = ofv.GenerarInput(rs(12).name,rs(12) ,"text",sForm,"30 ","dt","")

        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")
        Fila = fila & ofv.genrow("","","","","","","",Celda,"si")





        celda=ofv.GenCelda("","","","","","","","&nbsp;","si") & ofv.GenCelda("","VT","","","","","","TOPE DE MONTO MINIMO:","si")
        sInput = ofv.GenerarInput(rs(8).name ,rs(8) ,"text",sForm,"10 ","dtn","")

        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")
        Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


        celda=ofv.GenCelda("","","","","","","","&nbsp;","si") & ofv.GenCelda("","VT","","","","","","TOMAR EVENTOS RAIZ/RELACIONADOS MENORES AL TOPE:","si")
        sInput = ofv.gencomboFX(sForm,RS(9).NAME,rs(9),scomK,scomD," ")

        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")
        Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



        celda=ofv.GenCelda("","","","","","","","&nbsp;","si") & ofv.GenCelda("","VT","","","","","","SUMARIZAR EVENTOS REPETITIVOS:","si")
        sInput = ofv.gencomboFX(sForm,RS(10).NAME,rs(10),scomK,scomD," ")

        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")
        Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    CELDA=ofv.gencelda("","VT COLSPAN=3","","","","","","&nbsp;","si")
    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")



    CELDA=ofv.gencelda("","TT COLSPAN=3","","","","","","<u>INCLUIR CAMPOS DE BUSQUEDA</u>","si")
    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")

    CELDA=ofv.gencelda("","VT COLSPAN=2","","","","","","&nbsp;","si")
    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")



    for i =0 to ubound(ACampos)
        Adato=split(ACampos(i),"#")
        celda=ofv.GenCelda("","","","","","","","&nbsp;","si") & ofv.GenCelda("","VT","","","","","",adato(0),"si")
        sInput = ofv.gencomboFX(sForm,adato(1),rs(adato(1)),scomK,scomD," ")

        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")
        Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
    next






    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>