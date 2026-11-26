<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  sfuncion=" onload=" & chr(34) & "window.document.title='Cor OnTime - Consulta de Tareas';" & chr(34) & " "


  sHTML = ofv.menuHeader("",sfuncion)



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


if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if


 sopcion = request.querystring("opcion")

 tsk = request.querystring("tsk")
 fsreal = request.querystring("fsreal")
 fereal = request.querystring("fereal")
 fsPLAN = request.querystring("fsPLAN")
 fePLAN = request.querystring("fePLAN")
 STSSIGN = request.querystring("STSSIGN")
 sPADV = request.querystring("PADV")

     sopcion = "VP"





    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 
    XFCURR  = XFSTART
    XFHOY   = XFSTART



  StrSql = "Select * From PROYtareas where id = " & tsk
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rs1.eof then
 




 

         dFecha = CDATE(CLNG(XFSTART) + clng(cdate("01-01-1900")))

      celdas = ofv.GenCelda("","TTT ","","","","","",sImagen & "&nbsp;<B>TAREA :  " & RS1(4) & "." & rs1(5) & "</B>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
   '''   FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

    sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

      filas = ""

      celdas = ofv.GenCelda("","vT    ","","","","","","RESPONSABLE:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

    if RS1(13) = "NO" then
       sInput = "<B><I>SIN AFECTAR</I></B>"
    ELSE
      strsql = "Select * From USUARIOS where ID = '" & rs1(13) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
      sInput = "<B>" & sniv & "</B>"
    END IF
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt  COLSPAN=3 STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

        if rs1(16) = 1 then sInput = "HORA"
        if rs1(16) = 2 then sInput = "DIA"
        if rs1(16) = 3 then sInput = "SEMANA"
        if rs1(16) = 4 then sInput = "MINUTO"

        if rs1(15) > 1 then sinput = sinput & "S"

      celdas = ofv.GenCelda("","vt  ","","30%","","","","DURACION:","si")
      celdas = celdas & ofv.GenCelda("","","center","","5%","","","&nbsp;","si")
      sInput = rs1(15) & "&nbsp;" & sInput

      celdas = celdas & ofv.GenCelda("","Vt","","65%","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


     celdas = ofv.GenCelda("","vt  ","","","","","","MODO DE ASIGNACION:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")



IF SESSION("PRCF") <> "S" THEN
   XXXX = OWKF.GETPRYCNFG()
END IF



   IF rs1(24) > 0 THEN 



      IF rs1(22) > 0 THEN
         XTDUR = rs1(22)
      ELSE
         XTDUR = rs1(15)
      END IF
   
      IF rs1(23) > 0 THEN
         XUDUR = rs1(23)
      ELSE
         XUDUR = rs1(16)
      END IF

      SELECT CASE RS1(16)

             CASE 4

                  XVDUR = rs1(15)

             CASE 1

                  XVDUR = rs1(15)

             CASE 2
                  XVDUR = rs1(15) * CLNG(SESSION("DPP"))


             CASE 3 
                  XVDUR = rs1(15) * CLNG(SESSION("SPP")) * CLNG(SESSION("DPP"))

      END SELECT

      SELECT CASE XUDUR

             CASE 4

                  XWDUR = XTDUR
             CASE 1

                  XWDUR = XTDUR

             CASE 2
                  XWDUR = XTDUR * CLNG(SESSION("DPP"))


             CASE 3 
                  XWDUR = XTDUR * CLNG(SESSION("SPP")) * CLNG(SESSION("DPP"))

      END SELECT

      XWVDUR = CINT(XVDUR / XWDUR * 100)


        if XUDUR = 1 then sInput = "HORA"
        if XUDUR = 2 then sInput = "DIA"
        if XUDUR = 3 then sInput = "SEMANA"
        if XUDUR = 4 then sInput = "MINUTO"

        if XTDUR > 1 then sinput = sinput & "S"


  ''    celdas = ofv.GenCelda("","vt  ","","30%","","","","&nbsp;","si")
  ''    celdas = celdas & ofv.GenCelda("","","center","","5%","","","&nbsp;","si")
      sInput = XTDUR & "&nbsp;" & sInput

  ''    celdas = celdas & ofv.GenCelda("","Vt STYLE='COLOR:#993333;' ","","65%","","","",sInput & "&nbsp;&nbsp;(AL&nbsp;" & XWVDUR & "%)","si")
   ''   filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  END IF


        if rs1(24) = 0 then sInput = "FULL-TIME"
        if rs1(24) = 1 then sInput = "PART-TIME&nbsp;(" & SINPUT & ")"

    celdas = celdas & ofv.GenCelda("","vt  COLSPAN=2","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt  COLSPAN=3  ","","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")


'''''''''''''''''''''''''''''
    celdas = ""
    celdas = celdas & ofv.GenCelda("xt1","TTt style='display:none;border-style:solid;border-width:1;border-color:#888888;' ","center","33%","","","","Planificacion","si")
    celdas = celdas & ofv.GenCelda("xt2","TTt style='display:block;border-style:solid;border-width:1;border-color:#888888;' ","center","33%","","","","Avance/Imputaciones","si")
    celdas = celdas & ofv.GenCelda("xt3","TTR style='display:block;' ","center","33%","","","","<A CLASS=cc onclick=""document.all.xt1.style.display='block';document.all.xt2.style.display='none';document.all.xt3.style.display='none';document.all.xt4.style.display='block';document.all.DX2.style.display='block';document.all.DX4.style.display='none';"" HREF='javascript:void(0);' >Planificacion</A>","si")
    celdas = celdas & ofv.GenCelda("xt4","TTR style='display:none;' ","center","33%","","","","<A CLASS=cc onclick=""document.all.xt1.style.display='none';document.all.xt2.style.display='block';document.all.xt3.style.display='block';document.all.xt4.style.display='none';document.all.DX2.style.display='none';document.all.DX4.style.display='block';"" HREF='javascript:void(0);' >Avance/Imputaciones</A>","si")
    celdas = celdas & ofv.GenCelda("","UT","center","33%","","","","&nbsp;","si")

    filas = ofv.GenRow("","","","","22","","",celdas,"si")
    sHTML = sHTML & ofv.GenTabla("","","","95%","","0","","0","0","","",filas,"si")

''''''''''''''''''''''''''''''''''''''''''''''''

    sHTML = sHTML & "<div id='DX2' style='display:none;"
    sHTML = sHTML & "border-style:solid;border-width:1;border-color:#888888;"
    sHTML = sHTML & "background-color:" & session("APLBGD") & ";width=95%;"
    sHTML = sHTML & "' >"




    filas = ""



   '   celdas = ofv.GenCelda("","vt  COLSPAN=3 STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","","","","","","&nbsp;","si")
   '   filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vT    ","","","","","","MODO DE INICIO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      if rs1(6) = 0 then sInput = "INDEPENDIENTE"
      if rs1(6) = 1 then sInput = "EN SECUENCIA"
      celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    if rs1(6) = 0 then     

      celdas = ofv.GenCelda("","vT    ","","","","","","DEPENDENCIA:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

        if rs1(7) = 0 then sInput = "INICIO MANUAL"
        if rs1(7) = 1 then sInput = "LISTA DE TAREAS"
        if rs1(7) = 2 then sInput = "OTROS PROYECTOS"
      celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    END IF

      celdas = ofv.GenCelda("","vT    ","","","","","","FECHA DE INICIO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
        if rs1(17) = 0 then sInput = "NORMAL"
        if rs1(17) = 1 then sInput = "PREDETERMINADA"
      celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    vfondo = "<img src=imagenes/corpnt3.gif border=0 width=7 >"







          

       dFecha = cdate(FSPLAN + clng(cdate("01-01-1900")))
       SINPUTn1 = dFecha

       dFecha = cdate(FEPLAN + clng(cdate("01-01-1900")))
       SINPUTn2 = dFecha

       dFecha = cdate(FSREAL + clng(cdate("01-01-1900")))
       SINPUTn3 = dFecha

       dFecha = cdate(FEREAL + clng(cdate("01-01-1900")))
       SINPUTn4 = dFecha


      celdas = ofv.GenCelda("","vT    ","","","","","","INICIO PLANIFICADO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut","","","","","",sInputN1,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vT    ","","","","","","FIN PLANIFICADO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut","","","","","",sInputN2,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","TTT  COLSPAN=3   STYLE='border-style:solid;border-width:1;border-color:#888888;' ","CENTER","","","","","OBSERVACIONES","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      nCol = 80

      if len(wdescri) = 0 then wdescri = "Sin Observaciones" 

      if not  isnull(rs1(27)) then 
         wdescri = rs1(27)
      end if 




    t3b = "<font class=at><textarea class=dt  style='background-color:transparent;padding-top:2px;' "
    t3b = t3b & "name=" & rs1(27).name & " cols=" & nCol & " rows=5 "
    t3b = t3b & " readonly >"
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    Celda = t3b



      celdas = ofv.GenCelda("","UT  COLSPAN=3    ","CENTER","","","","",Celda,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

    sHTML = sHTML & "</div>"


'''''''''''''''''''''''''''''

    sHTML = sHTML & "<div id='DX4' style='display:block;"
    sHTML = sHTML & "border-style:solid;border-width:1;border-color:#888888;"
    sHTML = sHTML & "background-color:" & session("APLBGD") & ";width=95%;"
    sHTML = sHTML & "' >"





    filas = ""




     sinput3 = "0%"

     PADV = 0 


   if rs1(18) = 0 then


      celdas = ofv.GenCelda("","vT    ","","","","","","ESTADO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#336699;FONT-WEIGHT:BOLD;' ","","","","","","NO INICIADA","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vT    ","","","","","","INICIO PROYECTADO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#336699;FONT-WEIGHT:BOLD;'","","","","","",sInputN3,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vT    ","","","","","","FIN PROYECTADO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#336699;FONT-WEIGHT:BOLD;' ","","","","","",sInputN4,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


   elseif rs1(18) = 2 then


          sinput3 = "100%"

      celdas = ofv.GenCelda("","vT  ","","","","","","ESTADO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#0000aa;FONT-WEIGHT:BOLD;' ","","","","","","FINALIZADA","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vT  ","","","","","","INICIO REAL:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#0000aa;FONT-WEIGHT:BOLD;' ","","","","","",sInputN3,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vT  ","","","","","","FIN REAL:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#0000aa;FONT-WEIGHT:BOLD;' ","","","","","",sInputN4,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


   else

      celdas = ofv.GenCelda("","vT  ","","","","","","ESTADO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#008000;FONT-WEIGHT:BOLD;' ","","","","","","EN CURSO","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      sinput3 = sPADV & "%"

      celdas = ofv.GenCelda("","vT  ","","","","","","AVANCE:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      IF STSSIGN = 0 THEN

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#008000;FONT-WEIGHT:BOLD;' ","","","","","",sinput3 & "&nbsp;DENTRO DEL PLAN","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
             xcolor = "#008000"

      ELSEIF STSSIGN = 1  THEN

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#0000BB;FONT-WEIGHT:BOLD;' ","","","","","",sinput3 & "&nbsp;DENTRO DEL PLAN (DEMORADA)","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
             xcolor = "#0000BB"


      ELSEIF STSSIGN = 2  THEN

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#993333;FONT-WEIGHT:BOLD;' ","","","","","",sinput3 & "&nbsp;-&nbsp;DESVIO DEL PLAN","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
             xcolor = "#993333"


      ELSE

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:#008000;FONT-WEIGHT:BOLD;' ","","","","","",sinput3 & "&nbsp;DENTRO DEL PLAN","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
             xcolor = "#008000"

      END IF 




      celdas = ofv.GenCelda("","vT  ","","","","","","INICIO REAL:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut STYLE='FONT-WEIGHT:BOLD;' ","","","","","",sInputN3,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vT  ","","","","","","FIN PROYECTADO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

      celdas = celdas & ofv.GenCelda("","ut STYLE='COLOR:" & xcolor & ";FONT-WEIGHT:BOLD;' ","","","","","",sInputN4,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")






   end if


    sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")


   if rs1(18) > 0 then

      sHTML = sHTML & "<br>"

      celdas = ofv.GenCelda("","TTT colspan=4 style='border-style:solid;border-width:1;border-color:#888888;' ","center","","","","","REGISTRO DE ESTADO DE AVANCE","si")

      filas = ofv.GenRow("","","","","10","","",celdas,"si")

          StrSql = "Select * From PROYTAREASADV where codigo = '" & RS1(1) & "' AND TSK = '" & RS1(3) & "' ORDER BY id DESC "
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then

                xTrColor = "#FFFFFF"

                celdas = ofv.GenCelda("","TTr","center","","","","","FECHA","si")
                celdas = celdas & ofv.GenCelda("","TTr","center","","","","","IMPUTADO","si")
                celdas = celdas & ofv.GenCelda("","TTr","center","","","","","PORCENTAJE","si")
                celdas = celdas & ofv.GenCelda("","TTr","center","","","","","RESPONSABLE","si")
                filas = filas & ofv.GenRow("","","","","20","","",celdas,"si")


             do while not rsx.eof

                dFecha = cdate(rsX(3) + clng(cdate("01-01-1900")))
                SINPUTn1 = dFecha

                if rs1(16) < 4 then sInput = " H"


                if rs1(16) = 4 then sInput = " m"

                celdas = ofv.GenCelda("","uT","","10%","","","",SINPUTn1,"si")
                celdas = celdas & ofv.GenCelda("","ut","right","15%","","","",rsX(4) & sInput & "&nbsp;","si")
                celdas = celdas & ofv.GenCelda("","ut","right","15%","","","",rsX(5) & "%" & "&nbsp;&nbsp;&nbsp;","si")


                 sInput = "<B><I>SIN AFECTAR</I></B>"

                 strsql = "Select * From USUARIOS where ID = '" & rsX(6) & "' "
                 set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                 if not rsniv.eof then sniv = rsniv(1)
                 call ofv.cerrarconsulta(rsniv)
                 sInput = "<B>" & sniv & "</B>"


                celdas = celdas & ofv.GenCelda("","uti","","60%","","","",sInput,"si")

                filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","20","","",celdas,"si")
    


	        if xTrColor = "#FFFFFF" then
	           xTrColor = "#F4F4F4" 
	        ELSE
	           xTrColor = "#FFFFFF" 
	        END IF




             rsx.movenext
             loop  

          else

             celdas = ofv.GenCelda("","uT colspan=4  ","center","","","","","SIN IMPUTACIONES A LA FECHA","si")
             filas  = filas & ofv.GenRow("","","","","30","","",celdas,"si")



          end if

          sHTML = sHTML & ofv.GenTabla("","aa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","98%","","0","","-1","-2","","",filas,"si")
          sHTML = sHTML & "<br>"
   end if




    sHTML = sHTML & "</div>"

 
  end if











  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>