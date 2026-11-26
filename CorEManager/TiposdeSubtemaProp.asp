<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  redim nivel(9)
  nivel(0) = "C"
  nivel(1) = "D"
  nivel(2) = "E"
  nivel(3) = "F"
  nivel(4) = "G"
  nivel(5) = "H"
  nivel(6) = "I"
  nivel(7) = "J"
  nivel(8) = "K"
  nivel(9) = "L"
  sImagenCampo = "ARCHIVO"
  CantRegMostrarx = cdbl(request.querystring("CantRegMostrar"))
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  wPerfil = request.querystring("wPerfil")
  wNivel = request.querystring("wNivel")

  if isnull(wNivel) then wNivel = "NO"
  if len(wNivel) = 0  then wNivel = "NO"
  

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_wPerfil=" & wPerfil

  sHTML = ofv.FormHeader(session("FormN"))

 
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar




  set cn = ofv.conectar(ofv.strconn2)


  Strsql = "Select * From Subtema where tipobjeto = '" & wperfil & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 
  posI = 0

    sForm = "form" & posi


 ' Encabezados 1

  'Encabezados
'''  Formulario = ofv.Encabezados("Tipo","Descripcion","Nivel","Icono","")

  '''   Fila = Formulario




    'Tipo
    sInput = rs(0) & " - " & rs(1)
    Celda = ofv.gencelda("","ttt colspan=2","","","","","",sInput,"si")

     Fila = ofv.genrow("","","","","","","",Celda,"si")

    'Nivel
     Celda = ofv.gencelda("","tt","","20%","","","","Nivel","si")
  
    'NIvel
    sInput =  rs(4) 
    Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

     Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Icono
    'Celda = ofv.gencelda("","tt","","","","","","Icono","si")

    'Icono
    'sImagenValor = rs(15)
    'sInput = ofv.GenerarCombo3(sForm,ofv.Getfiles("Iconos"),sImagenCampo,sImagenValor)
    'Celda = Celda & ofv.gencelda("","vti ","","","","","",sInput,"si")

    'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Imagen
     'Celda = ofv.gencelda("","tt","","","","","","Imagen","si")

    'sImage = ofv.GenImage("Imagen","","","","Iconos/" & sImagenValor,"","")
    'Celda = Celda & ofv.gencelda("","fti ","","","","","",sImage,"si")

    'Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Subtema&volver=" & vuelta
    end if

    Fila = ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","80% ","","0","","-1","2","","",Fila,"si") & ""


 ' Encabezados 2

Formulario = ""

    sForm = "formB" & posi


    'Acepta niveles
     Celda = ofv.gencelda("","ttt colspan=10","","","","","","Acepta niveles","si")

     Fila = ofv.genrow("","","","","","","",Celda,"si")


Formulario = ""


    Celda =  ""



    'Acepta linea 1

    for ii = 0 to 4

        sinput = "<a  href=TiposdeSubtemaProp.asp?wperfil=" & rs(0) & "&CantRegMover=" & CantRegMoverx
        sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx  & "&wNivel=" & nivel(ii) & " ><b>"
        sinput = sinput  & "<font color=" & Session("CMFPNITFT") & ">" & nivel(ii) & "</font></b></a>" 
    
        Celda = Celda & ofv.gencelda("","cc","center","5%","","","",sinput,"si")

        sInput = ofv.combosino(sForm,rs(ii + 5),rs(ii + 5).name)

        if  rs(ii + 5) then
            sInput2 = "1"
        else
            sInput2 = "0"
        end if


        if  rs(ii + 5) then
            scolor = " bgcolor=#008000 "

        else
            scolor = "  bgcolor=#f7eedd  "

        end if

        Celda = Celda & ofv.gencelda("","ut","","10%","","","",sInput,"si")


    next 

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")



    Celda = ""

    'Acepta linea 2

    for ii = 5 to 9

        sinput = "<a  href=TiposdeSubtemaProp.asp?wperfil=" & rs(0) & "&CantRegMover=" & CantRegMoverx
        sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx  & "&wNivel=" & nivel(ii) & " ><b>"
        sinput = sinput  & "<font color=" & Session("CMFPNITFT") & ">" & nivel(ii) & "</font></b></a>"  
    
        Celda = Celda & ofv.gencelda("","cc","center","5%","","","",sinput,"si")

        sInput = ofv.combosino(sForm,rs(ii + 5),rs(ii + 5).name)
        sInput = replace(sInput,"select","select style='display:none;' ")

        if  rs(ii + 5) then
            sInput2 = "1"
        else
            sInput2 = "0"
        end if



        if  rs(ii + 5) then
            scolor = " bgcolor=#008000 "

        else
            scolor = "  bgcolor=#f7eedd  "

        end if

        Celda = Celda & ofv.gencelda("","ut","","10%","","","",sInput,"si")


    next 

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","aaa colspan=11","","","","","","&nbsp;","si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Subtema&volver=" & vuelta
    end if
    Fila =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa","","80% style='border-radius: 0px 0px 10px 10px'","","0","","-1","2","","",Fila,"si")


    posi = posi + 1
  end if

  call ofv.cerrarconsulta(rs)

  if wNivel <> "NO" then

     xGenConti1 = "Tipos de esquemas con nivel: " & wNivel
     xGenConti2 = "Tipo"
     xGenConti3 = "Descripcion"



    '1
     Celda = ofv.gencelda("","ttt colspan=2 ","","","","","",xGenConti1,"si")

     Fila = ofv.genrow("","","","","","","",Celda,"si")

    '2
     Celda2 = ofv.GenCelda("","tt","","20%","","","",xGenConti2,"si")

    '3
     Celda2 = Celda2 & ofv.gencelda("","tt","","80%","","","",xGenConti3,"si")

     Fila = fila & ofv.genrow("","","","","","","",Celda2,"si")

     Strsql = "Select * From Tipobjeto where nivel = '" & wNivel & "' order by descri "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then 
 
        do until rs.eof 
           'Tipo
           sInput = rs(0)
           Celda =  ofv.gencelda("","cc","","","","","",sInput,"si")

          'Descripcion
           sInput = rs(1)
           Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")
           Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

           rs.movenext
        loop 

    else

           sInput = "No hay esquemas para este nivel."
           Celda = ofv.gencelda("","cc colspan=2 ","","","","","",sInput,"si")
           Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    end if

    call ofv.cerrarconsulta(rs)

    sHTML = sHTML & ofv.GenTabla("","","","50% ","","0","","0","2","","",Fila,"si")



  end if

  Agregar = ""
  Buscar  = ""

  sAccion = "TiposdeSubtema.asp?CantRegMover=" & CantRegMoverx
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)




  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>