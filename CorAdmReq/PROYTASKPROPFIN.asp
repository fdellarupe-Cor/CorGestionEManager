<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.MenuHeader("","")



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






  XTSKSTS = RS(18)

  IF XTSKSTS = 2 THEN
     bModificar = 0
  END IF




      celdas = ofv.GenCelda("","cc colspan=4","","","","","",sImagen & "&nbsp;<BIG><B>FINALIZAR LA TAREA</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vT  ","","","","","","TAREA:","si")
      celdas = celdas & ofv.GenCelda("","vt","RIGHT","","","","",rs(4) & ".","si")
      sInput = rs(5)
      celdas = celdas & ofv.GenCelda("","vt  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vT  ","","","","","","ESTADO:","si")

      if  rs(18) = 0 then
          sinput1 = sImagenE0
          sinput2 = "NO INICIADA"

      elseif rs(18) = 2 then

          sinput1 = sImagenE2
          sinput2 = "FINALIZADA"

      else

          sinput1 = sImagenE1
          sinput2 = "EN CURSO"

      END IF


      celdas = celdas & ofv.GenCelda("","vt","","","","","",sinput1,"si")

      celdas = celdas & ofv.GenCelda("","vt  COLSPAN=2","","","","","",sinput2,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt  COLSPAN=4 STYLE='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#f7aa22;' ","","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt  COLSPAN=4  bgcolor=#ffffff STYLE='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#f7aa22;' ","CENTER","","","","","DURACION PLANIFICADA","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



      celdas = ofv.GenCelda("","vT  ","","","","","","TIEMPO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","",sImagendu,"si")
      sInput = rs(15)
      celdas = celdas & ofv.GenCelda("","vt  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vt  ","","","","","","UNIDAD DE TIEMPO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","",sImagenut,"si")

        if rs(16) = 1 then sInput = "HORAS"
        if rs(16) = 2 then sInput = "DIAS"
        if rs(16) = 3 then sInput = "SEMANAS"
        if rs(16) = 4 then sInput = "MINUTOS"

    celdas = celdas & ofv.GenCelda("","vt  COLSPAN=2","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt  COLSPAN=4 STYLE='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#f7aa22;' ","","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt  ","","","","","","MODO DE ASIGNACION:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","",sImagenut,"si")

        if rs(24) = 0 then sInput = "FULL-TIME"
        if rs(24) = 1 then sInput = "PART-TIME"


    celdas = celdas & ofv.GenCelda("","vt  COLSPAN=2","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt  COLSPAN=4 STYLE='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#f7aa22;' ","","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      formulario0 = ofv.GenForm("","","","","post","",filas,"si")


    filas = ofv.GenRow("","","","","","","",formulario0,"si")

    Celda = ofv.GenCelda("","abc colspan=4","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")




    FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")


''''''''

        USRTMP     = RS(15)
        USRTMPU    = RS(16)

        SELECT CASE USRTMPU
               CASE 1
                    xagrd = USRTMP

               CASE 4
                    xagrd = USRTMP

               CASE 2

                    xagrd = (USRTMP * CLNG(SESSION("DPP")))
 

               CASE ELSE

                    xagrd = (USRTMP * CLNG(SESSION("DPP")) * CLNG(SESSION("SPP")) )

        END select       

      if  rs(18) = 0 then

          XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900"))

      ELSE

          XFSTART = "NO"

      END IF


     celdas = ofv.GenCelda("","AA ","CENTER","5%","","","","&nbsp;","si")


     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYTASKPROP.asp"
     sinput = sinput  & " ><b>CANCELAR</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","LEFT","","","","",sInput,"si")

     celdas = celdas & ofv.GenCelda("","AA ","CENTER","5%","","","","&nbsp;","si")


     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYTASKADV.asp?OPCION=ATT&TADV=" & xagrd & "&PADV=100&FSTART=" & XFSTART
     sinput = sinput  & " ><b>CONFIRMA</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","LEFT","","","","",sInput,"si")


     fila = ofv.GenRow("","","","","","","",celdas,"si")

     FilaS3 = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")






    Celda = ofv.GenCelda("","","","","","","",FilaS1 & FILAS2 & FILAS3,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=#dddddd,direction=125);' ","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"





 
  IF session("ModMain2") = "CL" THEN
     sinput = "<a  class=tt href=Clientes.asp"
     sinput = sinput  & " ><b>&nbsp;Volver&nbsp;a&nbsp;Cliente&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

  ELSE
     sinput = "&nbsp;"
     celdas =  ofv.GenCelda("","","","1","","","",sInput,"si")
  END IF


  celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
  filas = ofv.GenRow("","","","","","","",celdas,"si")
  sHTML = sHTML  & ofv.GenTabla("","aa ","","98%","","0","","0","0","","",filas,"si")




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML






%>