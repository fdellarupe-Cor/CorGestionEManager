<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  xcm1A  = "0" & CHR(9) & "1" & CHR(9) & "2"
  xcm1B  = "CENTRADA" & CHR(9) & "IZQUIERDA" & CHR(9) & "DERECHA"

  xcmccA = ""
  xcmccB = ""

  xsep = ""

  for iiy=1 to 10
 
      xcmccA = xcmccA & xsep & iiy
      xcmccB = xcmccB & xsep & iiy & "&nbsp;Col"

      xsep   = CHR(9)

  next

set ofr = server.createobject("corscr.Corrender")

ofr.verfonts

  sxopc1  = ""
  sxopc1A = ""
  sxopc1C = ""
for ii=0 to ofr.wscreen.fontcount -1

  sxopc1  = sxopc1 & "<option value=" & chr(34) & ofr.wscreen.fonts(ii) & chr(34) & " >" & replace(ofr.wscreen.fonts(ii)," ","&nbsp;") & "</option> "
  sxopc1A = sxopc1A & sxopc1C & ucase(ofr.wscreen.fonts(ii))
  sxopc1C = chr(9) 
next



  sxopc2 = "<option value=8pt>8pt</option>"
  sxopc2 = sxopc2 & "<option value=9pt>9pt</option>"
  sxopc2 = sxopc2 & "<option value=10pt>10pt</option>"
  sxopc2 = sxopc2 & "<option value=12pt>12pt</option>"
  sxopc2 = sxopc2 & "<option value=14pt>14pt</option>"
  sxopc2 = sxopc2 & "<option value=16pt>16pt</option>"
  sxopc2 = sxopc2 & "<option value=18pt>18pt</option>"
  sxopc2 = sxopc2 & "<option value=20pt>20pt</option>"

  sxopc2A = ""
  sxopc2C = ""
for ii=1 to 14
  sxopc2A = sxopc2A & sxopc2C & cstr(ii + 7) & "PT" 
  sxopc2C = chr(9) 
next


  set cn  = ofv.conectar(ofv.strconn2)
  sxopc3  = ""
  sxopc3A = ""
  sxopc3B = ""
  sxopc3C = ""
  ii=0
  redim preserve ccol(0)
  redim preserve dcol(0)
  strsql = "Select * From Syscolor"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     do until rs.eof
       redim preserve ccol(ii)
       redim preserve dcol(ii)
       ccol(ii) = rs(2)
       dcol(ii) = rs(1)
       sxopc3   = sxopc3 & "<option value=" & chr(34) & rs(2) & chr(34) & ">" & rs(1)
       sxopc3   = sxopc3 & "</option> "
       sxopc3A  = sxopc3A & sxopc3C & rs(2)
       sxopc3B  = sxopc3B & sxopc3C & rs(1)
       sxopc3C  = chr(9) 
       ii = ii + 1
     rs.movenext
     loop
 end if
  call ofv.cerrarconsulta(rs)

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  sHTML = ofv.FormHeader(session("FormN"))




  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

  wperfil = request.querystring("wperfil")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_wperfil=" & wperfil

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From TABPERFIL where tabperfil = '" & wperfil & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 
  posI = 0




    sForm = "form" & posi


'''''''''''''''''''''

    'Descripcion

    sInput = rs(0)
    nperfil = rs(0)




   'Bordes
    Celda = ofv.gencelda("","tt","","20%","","","","Bordes","si")

    sInput = ofv.combosino(SForm,rs(2),rs(2).name)
    Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

   'Columnas
    Celda = ofv.gencelda("","tt","","","","","","Columnas","si")

 ''   sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",SForm,"5","dtn","si")

    sInput = ofv.gencomboFX(sForm,rs(7).name,rs(7),xcmccA,xcmccB,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sInput = "Bloques"
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Color
    Celda = ofv.gencelda("","tt","","","","","","Color","si")

    sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3),sxopc3A,sxopc3B,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Fondo
    Celda = ofv.gencelda("","tt","","","","","","Fondo","si")

    sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),sxopc3A,sxopc3B,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Font
    Celda = ofv.gencelda("","tt","","","","","","Fuente","si")

    sInput = ofv.gencomboFX(sForm,rs(5).name,rs(5),sxopc1A,sxopc1A,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Tamaño
    Celda = ofv.gencelda("","tt","","","","","","Tamaño","si")

    sInput = ofv.gencomboFX(sForm,rs(6).name,rs(6),sxopc2A,sxopc2A,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sInput = "Titulos de Columnas"
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Color
    Celda = ofv.gencelda("","tt","","","","","","Color","si")

    sInput = ofv.gencomboFX(sForm,rs(8).name,rs(8),sxopc3A,sxopc3B,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Fondo
    Celda = ofv.gencelda("","tt","","","","","","Fondo","si")

    sInput = ofv.gencomboFX(sForm,rs(9).name,rs(9),sxopc3A,sxopc3B,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Font
    Celda = ofv.gencelda("","tt","","","","","","Fuente","si")

    sInput = ofv.gencomboFX(sForm,rs(10).name,rs(10),sxopc1A,sxopc1A,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Tamaño
    Celda = ofv.gencelda("","tt","","","","","","Tamaño","si")

    sInput = ofv.gencomboFX(sForm,rs(11).name,rs(11),sxopc2A,sxopc2A,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sInput = "Contenido de Columnas"
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Color
    Celda = ofv.gencelda("","tt","","","","","","Color","si")

    sInput = ofv.gencomboFX(sForm,rs(12).name,rs(12),sxopc3A,sxopc3B,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Fondo
    Celda = ofv.gencelda("","tt","","","","","","Fondo","si")

    sInput = ofv.gencomboFX(sForm,rs(13).name,rs(13),sxopc3A,sxopc3B,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Font
    Celda = ofv.gencelda("","tt","","","","","","Fuente","si")

    sInput = ofv.gencomboFX(sForm,rs(14).name,rs(14),sxopc1A,sxopc1A,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'Tamaño
    Celda = ofv.gencelda("","tt","","","","","","Tamaño","si")

    sInput = ofv.gencomboFX(sForm,rs(15).name,rs(15),sxopc2A,sxopc2A,"")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname="
      sAccion = sAccion & ofv.convertircar(rs(0)," ","*") & "&tabla=tabperfil&volver=" & vuelta
    end if
    Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    PFtabla1 = ofv.GenTabla("","","","80% ","","0","","-1","2","","",Formulario,"si")





'''''''''''''''''''''''''

    ncanti = "0" & rs(7)

    ncanti = cint(ncanti)

    if ncanti < 1  then ncanti = 1

    if ncanti > 10 then ncanti = 10


    sForm = "formB" & posi

    sInput = "Configuracion de Columnas"
    Celda = ofv.gencelda("","TTT COLSPAN=3 ","","","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")


   'Bordes
    Celda = ""
    Celda = Celda & ofv.gencelda("","tt","","20%","","","","Columna","si")
    Celda = Celda & ofv.gencelda("","tt","","40%","","","","Ancho","si")
    Celda = Celda & ofv.gencelda("","tt","","40%","","","","Alineacion","si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    FOR II=1 TO 10

        IF II > NCANTI THEN EXIT FOR

        Celda = ""

        sInput = II
        Celda = Celda & ofv.gencelda("","cc","center","","","","",sInput,"si")

        sInput = ofv.GenerarInput(rs(15 + II).name,rs(15 + II),"text",SForm,"5","dtn","si")
        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")

        sInput = ofv.gencomboFX(sForm,rs(25 + II).name,rs(25 + II),xcm1A,xcm1B,"")
        Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")

        Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    NEXT




    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname="
      sAccion = sAccion & ofv.convertircar(rs(0)," ","*") & "&tabla=tabperfil&volver=" & vuelta
    end if
    Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    PFtabla2 = ofv.GenTabla("","TBP","","100% ","","0","","-1","2","","",Formulario,"si")


  end if

    sInput = rs(0) & " - " & rs(1)
    Celda  = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")

    Fila   = ofv.genrow("","","","","","","",Celda,"si")

    Celda  = ""
    sInput = PFtabla1
    Celda  = Celda & ofv.gencelda("","UT valign=top","","40%","","","",sInput,"si")

    sInput = PFtabla2
    Celda  = Celda & ofv.gencelda("","UT valign=top","center","60%","","","",sInput,"si")


    Fila   = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sHTML  = sHTML & ofv.GenTabla("","aaa","","80% ","","0","","-1","2","","",Fila,"si") & "<br>"



  Agregar = ""


  'Boton test
  sAccion = "PerfilesdeTablasTest.asp?nperfil=" & nperfil
  Buscar = ofv.BotonTest(sAccion)


  sAccion = "PerfilesdeTablas.asp?Cantregmover=" & Cantregmoverx
  Celda = ofv.BotonVolver(sAccion)



  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no","",Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>