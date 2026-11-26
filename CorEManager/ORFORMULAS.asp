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
  SCODINDI       = request.querystring("ID")


  if session("SCODINDI")="" then
   session("SCODINDI")=SCODINDI
  else
   SCODINDI=session("SCODINDI")
  end if



  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

'''''''''''''''''''''''''''''''''''''''
'' INCORPORO EVENTO EN BASE DE EVENTOS DE PERDIDA

  if sopcion = "SI" then

 sfechai = clng(date()) - clng(cdate("01-01-1900"))

        strsql = "Insert into ORFORMULAS "
        strsql = strsql & " (DEFINICION,CANTVAR,TIPOF,DESCRI,FECHA,ESTADO"
        strsql = strsql & " ) Values ('UNICO',	1,1,'',"
        strsql = strsql    & sfechai
        strsql = strsql & ",1)"
       '  response.write strsql & "<br>"
         xxx=SqlExLog("INSERT",strsql,CN,"ORFORMULAS","","","PAR")
       ' call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if


  if sOPCION = "ELIM" then



       ' strsql = "DELETE FROM ORINDICADORES WHERE OBJID = 0" & SCODINDI
         '''response.write strsql & "<br>"
      '  call ofv.crearconsultaEx(StrSql,cn,1,volver)


  END IF


  ''ARMO COMBO DE TIPO DE ESTADOS


    scomD = "MANUAL" & chr(9) & "AUTOMATICO" & CHR(9) & "SEMIAUTOMATICO"
    scomK = "1" & chr(9) & "2" & chr(9) & "3"



'

   Strsql = "select * FROM orformulas"
   Celda = ""

   wkproccod = "&nbsp;"

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","FORMULAS","si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")




  'Encabezados
  Formulario = ""

    Celda = ""


    Formulario = ""




  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

    Celda = ofv.gencelda("","vT  style='border-style:SOLID;border-width:1;border-color:#f93333;' ","center","","","","","<big>NO HAY CARGADAS FORMULAS</big>","si")
    Formulario = ofv.genrow("","","","","60","","",Celda,"si")

  ELSE

    Celda = ofv.gencelda("","VTI COLSPAN=2 STYLE='background-color:#bbbbbb;color:#ffffff;' ","","1%","","","","","si")
    'REGISTRO
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","DESCRIPCION","si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","FORMULA","si")

    'Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","CANT. VARIABLES","si")



    'FECHA CREACION
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","FECHA CREACION","si")
  '  Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;'","","","","","","VARIABLES","si")

    Formulario = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  posI = 0

  Strsql = "Select * From ORTIPOFORMULA ORDER BY DESCRI"


  scomD = "1" & chr(9) & "2" & CHR(9) & "3"
  scomK = "1" & chr(9) & "2" & chr(9) & "3"

  WKSETTABLA="ORFORMULAS"
  if not rs.eof then rs.move(CantRegMoverx)
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    Celda = ""

    'Eliminar
	if bSuprimir and rs("estado") <> 0 then
	  sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
	  sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
	  Celda = ofv.BotonEliminar(rs(2),sHref)
	else
	  Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
	end if


    'ID
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")

    'descri
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    IF rs("estado") <> 0 THEN
    sInput = ofv.GenerarCombo(sForm,rs(1).name,rs(1).value,Strsql,cn,0,1,"")
    ELSE
     sInput =RS("DESCRI")
    END IF

    Celda = Celda & ofv.gencelda("","ut ","","","","","",sInput,"si")


    '



    ' sInput = ofv.GenerarCombo(sForm,rs(1).name,rs(1).value,Strsql,cn,0,1,"")
    ' Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")



    'FECHA


    sInput = GTFechadorSH(RS("FECHA"))
	Celda = Celda & ofv.gencelda("","ut ","","","","","",sInput,"si")



    Filax  = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=ORFORMULAS&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")


    posi = posi + 1
    rs.Movenext
  Loop
  END IF
  sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar  then
    sAccion = "ORFORMULAS.asp?opcion=SI&ID=" & SCODINDI
    Agregar = ofv.BotonAgregar(sAccion)
  end if
  Buscar  = ""
  Celda   = ""


  'consulta
  sAccion = "ORFORMULASCons.asp?" & request.querystring
  Celda   = ofv.BotonConsultar(sAccion)


  str1 = "Select descri FROM orformulas "    ' orden pagina
  str2 = "Select descri FROM orformulas order by descri"    ' orden alfabetico
 Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))


 ''' sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
   sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Buscar,Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)




  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>