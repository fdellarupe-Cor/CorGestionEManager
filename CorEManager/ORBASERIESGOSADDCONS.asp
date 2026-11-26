<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<!-- #INCLUDE FILE="IncFile/reportes.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  sCODTRACER   = request.querystring("CODTRACER")
  sOPCSET      = request.querystring("OPCSET")



'response.write request.querystring
''''  sHTML = ofv.FormHeader(session("FormN"))


    Celda = ""
    APLCHECK()


    set oPrtG = new ClsReportes 


    Set oPrtG.Ofv = Ofv


   sHTML = corstyler()

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
   xsHTML=  oPrtG.RepHeader2(cstr(""), cstr("600"))



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)
  set cn3 = ofv.conectar(ofv.strconn2)


'''''''''''''''''''''''''''''''''''''''


    Filax  = ""

    Formulario = ""

          WSPSEN1 = "<FONT COLOR='#99CCEE' SIZE=4>&nbsp;&nbsp;&#0149;</FONT>"
          WSPSEN2 = "<FONT COLOR='#880000' SIZE=4>&nbsp;&nbsp;&#0149;</FONT>"





  Strsql = "  select * FROM ORTRACER WHERE CODTRACER = '" & SCODTRACER & "' "

 'response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''



    XORWKESRAIZ = "NO"
    XORWKESRAIZDSC = "&nbsp;"
    XORWKENCADSC = "&nbsp;"
    XORWKENCA = "NO"
    XORWKENCACOD = "&nbsp;"

    SINPUT = "PATRON DE DATOS OBLIGATORIOS - "


    IF SOPCSET = "DEF" THEN
       SINPUT = SINPUT & "DEFAULT"
    ELSE
       sInputX = "---"
       strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & SOPCSET
      ' response.write strsql
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       SINPUT = SINPUT & "ESPECIAL " & sInputx
    END IF



    SINPUT2 = "&nbsp;"
    WSPCOLOR = "#000000"

    Celda = ofv.gencelda("","VTI COLSPAN=6 STYLE='background-color:" & WSPCOLOR & ";color:#ffffff;'  ","","","","","",SINPUT,"si")
    Filax  = Filax & ofv.genrow("","","","","30","","",Celda,"si")


      sInputD1 = ""
      sInputD2 = ""




    Celda = ""



    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","F.CARGA:&nbsp;","si")

    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(2))
    end if
    SINPUT = dFechan & sInputD1 & sInputD2
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")



    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","ESTADO:&nbsp;","si")

    'ESTADO

     wkestadoD = "<font color='#000000' >PATRON DE DATOS</font>"
    sInput = wkestadoD
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","F.INICIO:&nbsp;","si")

       IF CSTR(RS(3) & "") = "0" THEN
           WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN
      Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")







    Celda = ""

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","F.DESCUBRIMIENTO:&nbsp;","si")


       IF CSTR(RS(4) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

      Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","F.R.CONTABLE:&nbsp;","si")


       IF CSTR(RS(5) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

      Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","F.FINALIZACION:&nbsp;","si")


       IF CSTR(RS(23) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

      Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")



    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


'''''''''''''''''''''''''''


    Celda = ""


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","PROCESO:&nbsp;","si")



       IF CSTR(RS(6) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","SUBPROCESO:&nbsp;","si")

       IF CSTR(RS(7) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN


    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","ACTIVIDAD:&nbsp;","si")

       IF CSTR(RS(9) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN



    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


'''''''''''''''''''''''''''










'''''''''''''''''''''''''''


    Celda = ""


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CARACTER:&nbsp;","si")

       IF CSTR(RS(24) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CUENTA&nbsp;CONTABLE:&nbsp;","si")

       IF CSTR(RS(29) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","NRO ASIENTO CONTABLE:&nbsp;","si")
    'ASIENTO CONTABLE

       IF CSTR(RS(41) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN


    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")






    Celda = ""


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","REF.&nbsp;CLIENTE:&nbsp;","si")

       IF CSTR(RS(28) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")








    Celda = ""


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","MN.BRUTO&nbsp;","si")


       IF CSTR(RS(12) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN


    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","COBERTURA SEGURO:&nbsp;%" ,"si")

       IF CSTR(RS(21) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","PRODUCTO/SERVICIO:&nbsp;","si")



       IF CSTR(RS(25) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Celda = ""

    if session("APLMDLOR") <> "BE"  THEN
     Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","AREA/SECTOR","si")
    ELSE
     Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","SUCURSAL:<BR><BR>AREA:&nbsp;","si")

    END IF


       IF CSTR(RS(20) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",  sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","	RIESGO&nbsp;ASOCIADO&nbsp;","si")

       IF CSTR(RS(26) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","AREA SOPORTE:&nbsp;","si")

       IF CSTR(RS(27) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")





'30
    Celda = ""
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","TIPO EVENTO:&nbsp;","si")
    sInput = WKTEVENTO

       IF CSTR(RS(30) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CATEGORIA:&nbsp;","si")
    sInput = WKTCATEGORIA

 
       IF CSTR(RS(31) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","EVENTO&nbsp;DE&nbsp;RIESGO:&nbsp;","si")



       IF CSTR(RS(10) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput ,"si")

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")





'30
    Celda = ""
    'OBSERVACIONES
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","DESCRIPCION:&nbsp;","si")

    sinput2 = "&nbsp;"

       IF CSTR(RS(11) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","LINEA NEGOCIOS:&nbsp;","si")

       IF CSTR(RS(8) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","COSTO OPORTUNIDAD:&nbsp;","si")


       IF CSTR(RS(39) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

       sInput = sInput & WSPSEN

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")







  END IF


    Formulario = FilaX



  sHTML = sHTML  & xshtml & "<CENTER>"
  sHTML = sHTML & "<BR>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:3;border-color:#888888;' ","","90% ","","0","","0","0","","",Formulario,"si")
  sHTML = sHTML & "<BR>" & tablab


 
  call ofv.cerrarconsulta(rs)


  call ofv.cerrarconn(cn)



  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML






%>