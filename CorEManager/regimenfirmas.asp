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
  SOPCION=REQUEST.QUERYSTRING("OPCION")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)
  sHTML = ofv.FormHeader(session("FormN"))

  IF SOPCION="SI" THEN

          strsql = "Insert into REGIMENFIRMAS "
          strsql = strsql & " (reflink,rellink,site)"
          strsql = strsql & " Values (NULL,0," & SESSION("CLIEMP") &")"
   '''''''RESPONSE.WRITE STRSQL
          xxx=SqlExLog("INSERT",strsql,CN,"REGIMENFIRMAS","","","")


  END IF



  'Encabezados

   Formulario = ofv.Encabezados("Id","DESCRIPCION","LINK&nbsp;A&nbsp;DOCUMENTO")


  Strsql = "Select * From REGIMENFIRMAS  where site=" & SESSION("CLIEMP") & " ORDER BY DESCRI"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=REGIMENFIRMAS&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(0),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'id
    sInput = rs(0)
    Celda = Celda & ofv.GenCelda("","cc","center","5","","","",sInput,"si")





    'Descripcion

    IF BMODIFICAR THEN
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"150","dtbs","")
    ELSE
    SINPUT=RS(1)
    END IF

    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    'Propiedades

     IF BMODIFICAR then
       vlk = vuelta
      wstextoLNK = rs(2) & ""

       if  isnumeric(wstextoLNK) and wstextoLNK <> "0" then
           sImagenP = "<font face='CorFlaticon' size=2 color=#005500 >&#0232;</font>"
       else
           sImagenP = "<font face='CorFlaticon' size=2 color=#f7aa22 >&#0232;</font>"
       end if



            sID = "EdGENLnk.asp?"
            sID = sID & "ID=" & rs(0) & "&TABLA=regimenfirmas&Codobjeto=" & RS(0)
            sID = sID & "&Tipobjeto=REGIFX" & sTipo & "&volver=" & vlk & "_Codobjeto=" & nCodigo
            sID = sID & "_Tipobjeto=" & sTipo

   '''  & "&nblk=AL"

         sHref = sID


          sinputk = "<a href='" & sHref
          sinputk = sinputk & "' title='Actualizar link a otro documento' ><i class='las la-arrow-right'></i></a>"
          slinkX   = "&nbsp;" & sinputk



     ELSE
       slinkX="&nbsp;"
     END IF
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",slinkX,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")



    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=REGIMENFIRMAS&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    posi = posi + 1
    rs.Movenext
  Loop

  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")


  retornar = ""



  'Botonera
  Agregar = ""
  if bAgregar then

 '''sAccion = "regimenfirmas.asp?Opcion=SI"

    sAccion = "AgregarRegimenFirmas.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=REGIMENFIRMAS"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From REGIMENFIRMAS where site=" & SESSION("CLIEMP") & "  ORDER BY DESCRI"
  str2 = "Select descri From REGIMENFIRMAS where site=" & SESSION("CLIEMP") & "  ORDER BY DESCRI"
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

     Listar  = ""

  sAccion = "REGIMENFIRMASCons.asp?" & request.querystring
  Celda = ""'ofv.BotonConsultar(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,Listar,celda & retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>