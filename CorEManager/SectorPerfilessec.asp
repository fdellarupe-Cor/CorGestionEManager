<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  seco = request.querystring("eco")

  if seco<> "" then
   session("seco")=seco
   session("Svuelta")=request.querystring
  else
   seco=session("seco")
  end if

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)



  sHTML = ofv.FormHeader(" ")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



  set cn = ofv.conectar(ofv.strconn2)











    posI = 0
    strsql="Select ECOPERFILDEP.*,eco.descri FROM ECOPERFILDEP "
    strsql=strsql & " left join eco on (ECOPERFILDEP.ptosector=eco.eco) "
    strsql =strsql & " where objid=" & seco
   ' response.write strsql
    set rs= ofv.crearconsultaEx(StrSql,cn,1,volver)

    Celda = ofv.GenCelda("","UTI" & wkstylo2,"","2%","","","","<B>SECTOR:&nbsp;&nbsp;" & rs("descri") & "</B>","si")

	   Fila  = Fila & ofv.genrow("","","","","25","","",Celda,"si")


    Formulario = Formulario & FILA

   IF Not rs.EOF THEN



    sForm = "formRB" & posi




    'Perfil
    Sql1 = "Select * From TIPOPERFIL WHERE TIPOPERFIL IN (SELECT CODPERFIL FROM ECOPERFIL WHERE CODSECTOR = 0" & rs(4) & ") "
    sql1 = Sql1 & " order by descri"
   ' response.write request.querystring
    sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5).value,sql1,cn,0,1,"")

    Celda = ofv.gencelda("","uti  ","","60%","","","",sInput,"si")

    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=EcoperfilDEP&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")





  END IF

  sAccion = "SectorPERFILESfun.asp?" & session("Svuelta")
  retornar = ofv.BotonVolver(sAccion)
  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","","","95%","","0","","1","1","","",Formulario,"si") & "</center>"

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","","no",agregar,retornar,"")

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>