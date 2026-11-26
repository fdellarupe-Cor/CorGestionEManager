<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FORMHeader("")



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



 sopcion = request.querystring("opcion")
 CTANUEVA = request.form("CTANUEVA")



 IF SOPCION = "SI" THEN

    XXXX = OWKF.GENTSKCUENTATRF(CTANUEVA)

    RESPONSE.REDIRECT "PROYTASKPROP.ASP"

 END IF 



    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#FFFFFF SIZE=2>l</FONT>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3 = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
    sImagen4s = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=2>&#0128;</FONT></B>"
    sImagen4O = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=2>&#0128;</FONT></B>"
    sImagen4n = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>x</FONT></B>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"
    sImagen6SS = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=2>&#0246;</FONT></B>"
    sImagen6N = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=3>3</FONT></B>"
    sImagen6OT = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=3>&#0254;</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B>"
    sImagenut = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=3>&#0192;</FONT></B>"
    sImagendu = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=3>h</FONT></B>"
    sImagenfi = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>R</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"
    sImagenE0 = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=4>6</FONT></B>"
    sImagenE1 = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"
    sImagenE2 = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=4>&#0251;</FONT></B>"

  sopcion = "VP"

  StrSql = "Select * From PROYTAREAS where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



  


     SINPUTF1 = ""
     SINPUTF2 = ""

     SFORM = "Usuarios1"


      celdas = ofv.GenCelda("","ttr colspan=4","","","","","","TRANSFERENCIA DE CUENTA DE PRESUPUESTO","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","UT  ","","","","","","TAREA:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","","",rs(4) & ".","si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",SFORM,"80","dt","")
      else
         sInput = rs(5)
      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","ttr  COLSPAN=4   ","","","","","","CONFIRMAR TRANSFERENCIA DE CUENTA","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","VT  COLSPAN=4   ","","","","","","DE CUENTA:","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      StrSql = "Select DESCRIPCION From PROYCUENTAS where ID = 0" & RS(32)
      set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsX2.eof then
         sInput  = rsX2(0)
      ELSE
         sInput  = "SIN ASIGNAR"
      END IF
      celdas = ofv.GenCelda("","UT  COLSPAN=4","CENTER","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","VT  COLSPAN=4   ","","","","","","A CUENTA:","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      StrSql = "Select DESCRIPCION From PROYCUENTAS where ID = 0" & CTANUEVA
      set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsX2.eof then
         sInput  = rsX2(0)
      ELSE
         sInput  = "SIN ASIGNAR"
      END IF
      celdas = ofv.GenCelda("","UT  COLSPAN=4","CENTER","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



      celdas = ofv.GenCelda("","UT  ","","","","","","&nbsp;","si")
      sInput = ofv.GenerarInput("CTANUEVA",CTANUEVA,"hidden",SFORM,"1","dt","")
      sInput = sInput & ofv.GenerarInput("boton","CONFIRMA","submit",SFORM,"10","bt","")
      celdas = celdas & ofv.GenCelda("","UT COLSPAN=2 ","center","","","","",sinput,"si")
      celdas = celdas & ofv.GenCelda("","UT","","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    if bModificar then
      sAccion = "PRYCUENTASTRF.asp?opcion=SI"
      formulario0 = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    sHTML = sHTML & ofv.GenTabla("","aa STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","80%","","0","","0","0","","",formulario0,"si")

 
    sHTML = sHTML & "</center>"




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""

  sAccion = "PROYTASKprop.asp"
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,Celda)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML






%>