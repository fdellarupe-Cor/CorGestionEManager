<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  dim ArrPorcentaje(1)
  dim arrVar(6)


  arrVar(0)="A"
  arrVar(1)="B"
  arrVar(2)="C"
  arrVar(3)="D"
  arrVar(4)="E"
  arrVar(5)="F"
  arrVar(6)="G"

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  opcion=request.querystring("opcion")
  varz=request.querystring("varz")
  bVOLVER=request.querystring("VOLVER")
  id=request.querystring("id")

' response.write REQUEST.QUERYSTRing
    IF id="" THEN
      ID=SESSION("ID")
    ELSE
     SESSION("ID")=ID
   END IF




  set cn = ofv.conectar(ofv.strconn2)
  sHTML = ofv.FormHeader(session("FormN"))


  select case opcion

  CASE "VERINFINDI"
    WKSETTABLA = "CSORVARIABLES2"


       Strsql = "Select * From CSORVARIABLES3  WHERE  IDINDICADOR=0" & id
       RESPONSE.WRITE STRSQL
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


       posI = 0
       strsql =" select a.* ,b.DEF ,b.definicion from orindicadores a,csoRformulas b where a.objid=0" & id
  	   strsql= strsql & " and a.formula=b.objid"
  	 '  response.write strsql

  	   set rsind = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  	   if not rsind.eof then
  	    UNIDADV=rsind(5)
  	    UNIDADRES=rsind(11)
  	    Strsql="select descri from ormedidas where objid=" & UNIDADV
  	     set rsun = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  	     UNIDADV=rsun(0)
  	     call ofv.cerrarconsulta(rsun)

  	     Strsql="select descri from ormedidas where objid=" & UNIDADRES
  	 	    set rsun = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  	 	    UNIDADRES=rsun(0)
  	     call ofv.cerrarconsulta(rsun)


  	    defFormula=rsind("def")

  	    Celda =  ofv.gencelda("","UTi COLSPAN=3 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","",RSIND("DESCRIPCION"),"si")
  	    Celda = CELDA & ofv.gencelda("","UTi COLSPAN=2 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","","FUNCION :      " & RSIND("DEF"),"si")
  	    filaZ= ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda 	 ,"si")
  	    CELDA=""
  	   end if

  	     Celda =  ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","","","si")
  	   	 Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","NOMBRE","si")
  	   	 Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","VALOR","si")
  	   	 Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","MEDIDA","si")
  	   	 Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","DESCRIPCION","si")
  	      Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","FECHA","si")
  	      Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","UMBRAL","si")

        filaZ=filaZ & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda 	 ,"si")
        CALL OFV.CERRARCONSULTA("RSIND")
       Do While (Not rs.EOF)
         sForm = "form" & posi


         IF RS("VARNOMBRE")="Z" THEN
         Celda=""
         Varz=rs(0)
        if not isnull(rs("fechact")) then
         sInput = "RESULTADO:"
         Celda = ofv.gencelda("","UTi COLSPAN=2 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","",sInput,"si")

        ' Celda = CELDA & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","","si")

         if defFormula="PORCENTAJE" THEN
          sInput =  replace(RS("VALOR")&"" , "," ,".") & "%"
         ELSE
          sInput =  replace(RS("VALOR")&"" , "," ,".")
         END IF

         Celda = CELDA&ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#FFBBBB;background-color:#FFBB99' ","","","","","",sInput,"si")

          SInput = UNIDADRES
          Celda = CELDA&ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#BBBBBB' ","","","","","",sInput,"si")


         sInput = GTFechadorSH(RS("FEChact"))


          Celda = CELDA & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","","si")
          Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput  ,"si")
         ' Strsql = "Select * From orumbrales  WHERE VALOR1 >=0"&RS("VALOR") &" AND IDINDICADOR="&id
         svalor = replace(RS("VALOR")&"" , "," ,".")
         Strsql = "Select * From orumbrales  WHERE VALOR1 <=0"&svalor& " AND VALOR2>=0"& svalor & " AND IDINDICADOR="&id
          'RESPONSE.WRITE STRSQL


           set rsum = ofv.crearconsultaEx(StrSql,cn,1,parametros)

            if rsum.eof then
              Celda = Celda &ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:#ff0000;' ","CENTER","5%","","","","<font color=#ffffff>FUERA DE RANGOS</font>","si")
            else
              Celda = Celda &ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:" & rsum("color") & ";' ","CENTER","5%","","","",rsum("descripcion"),"si")
            end if
            call ofv.cerrarconsulta("rsum")

          ELSE
           Celda = Celda &ofv.gencelda("","UT COLSPAN=7 STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:#ff0000;' ","CENTER","5%","","","","<font color=#ffffff>NO SE CONFIRMO VARIABLES PARA ESTE PERIODO</font>","si")
       end if


         ELSE


         sInput = rs(5) & "&nbsp;"
         Celda = ofv.gencelda("","UTi ALIGN=RIGHT  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")



         sInput =RS(9)& "&nbsp;"
         Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


          sInput =rs(6) ' ofv.GenerarInput(rs(6).name,rs(6),"text",sForm,"30","dtn","")
          Celda = Celda & ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


         sInput = UNIDADV
         Celda = CELDA&ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#BBBBBB' ","","","","","",sInput,"si")

          SInput =RS(4)& "&nbsp;"
          Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


          sInput = GTFechadorSH(RS("FEChact"))
          sfechai = clng(date()) - clng(cdate("01-01-1900"))

     	    sInputf = ofv.GenerarInput("fechact",sfechai,"hidden	",sForm,"70","dt","")






          Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput & sinputf,"si")


          END IF


          Fila  =  ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda 	 ,"si")



         xFila = ofv.genrow("","","","","","","",Celda,"si")


         Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

         posi = posi + 1
         rs.Movenext
       Loop



      sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","", FILAZ & Formulario,"si")



  case "VERINFO"




   WKSETTABLA = "ORVARIABLES"


     Strsql = "Select * From ORVARIABLES  WHERE  IDINDICADOR=0" & id
     'RESPONSE.WRITE STRSQL
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


     posI = 0
     strsql =" select a.* ,b.descri AS DEFF ,b.definicion from orindicadores a,csoRformulas b where a.objid=0" & id
	   strsql= strsql & " and a.formula=b.objid"

	   set rsind = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	   if not rsind.eof then
	    UNIDADV=rsind(5)
	    UNIDADRES=rsind(11)
	    Strsql="select descri from ormedidas where objid=" & UNIDADV
	     set rsun = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	     UNIDADV=rsun(0)
	     call ofv.cerrarconsulta(rsun)

	     Strsql="select descri from ormedidas where objid=" & UNIDADRES
	 	    set rsun = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	    UNIDADRES=rsun(0)
	     call ofv.cerrarconsulta(rsun)


	    defFormula=rsind("definicion")

	    Celda =  ofv.gencelda("","UTi COLSPAN=3 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","",RSIND("DESCRIPCION"),"si")
	    Celda = CELDA & ofv.gencelda("","UTi COLSPAN=2 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","","FUNCION :      " & RSIND("DEFF"),"si")
	    filaZ= ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda 	 ,"si")
	    CELDA=""
	   end if

	     Celda =  ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","","","si")
	   	 Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","NOMBRE","si")
	   	 Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","VALOR","si")
	   	 Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","MEDIDA","si")
	   	 Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","DESCRIPCION","si")
	      Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","FECHA","si")
	      Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","UMBRAL","si")

      filaZ=filaZ & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda 	 ,"si")
      CALL OFV.CERRARCONSULTA("RSIND")
     Do While (Not rs.EOF)
       sForm = "form" & posi


       IF RS("VARNOMBRE")="Z" THEN
       Celda=""
       Varz=rs(0)
      if not isnull(rs("fechact")) then
       sInput = "RESULTADO:"
       Celda = ofv.gencelda("","UTi COLSPAN=2 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","",sInput,"si")

      ' Celda = CELDA & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","","si")

       if defFormula="PORCENTAJE" THEN
        sInput = replace(RS("VALOR")&"" , "," ,".") & "%"
       ELSE
        sInput =replace(RS("VALOR")&"" , "," ,".")
       END IF

       Celda = CELDA&ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#FFBBBB;background-color:#FFBB99' ","","","","","",sInput,"si")

        SInput = UNIDADRES
        Celda = CELDA&ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#BBBBBB' ","","","","","",sInput,"si")


       sInput = GTFechadorSH(RS("FEChact"))


        Celda = CELDA & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","","si")
        Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput  ,"si")
        ' Strsql = "Select * From orumbrales  WHERE VALOR1 >=0"&RS("VALOR")& " AND IDINDICADOR="&id
        Svalor=replace(RS("VALOR")&"" , "," ,".")
        Strsql = "Select * From orumbrales  WHERE VALOR1 <=0"& Svalor& " AND VALOR2>=0"& Svalor & " AND IDINDICADOR="&id

       ' response.write strsql
         set rsum = ofv.crearconsultaEx(StrSql,cn,1,parametros)

          if rsum.eof then
            Celda = Celda &ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:#ff0000;' ","CENTER","5%","","","","<font color=#ffffff>FUERA DE RANGOS</FONT>","si")
          else
            Celda = Celda &ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:" & rsum("color") & ";' ","CENTER","5%","","","",rsum("descripcion"),"si")
          end if
          call ofv.cerrarconsulta("rsum")
     end if


       ELSE


       sInput = rs(5) & "&nbsp;"
       Celda = ofv.gencelda("","UTi ALIGN=RIGHT  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")



       sInput =RS(9)& "&nbsp;"
       Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


        sInput =replace(RS(6)&"" , "," ,".") ' ofv.GenerarInput(rs(6).name,rs(6),"text",sForm,"30","dtn","")
        Celda = Celda & ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


       sInput = UNIDADV
       Celda = CELDA&ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#BBBBBB' ","","","","","",sInput,"si")

        SInput =RS(4)& "&nbsp;"
        Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


        sInput = GTFechadorSH(RS("FEChact"))
        sfechai = clng(date()) - clng(cdate("01-01-1900"))

   	    sInputf = ofv.GenerarInput("fechact",sfechai,"hidden	",sForm,"70","dt","")






        Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput & sinputf,"si")


        END IF


        Fila  =  ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda 	 ,"si")



       xFila = ofv.genrow("","","","","","","",Celda,"si")


       Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

       posi = posi + 1
       rs.Movenext
     Loop



    sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","", FILAZ & Formulario,"si")




  CASE ELSE



 ' ME FIJO SI NO TIENE NADA PENDIENTE

  strsql=" select * from csorvariables2 where  activa=0 and idindicador=0" & id


  set rsvariables =ofv.crearconsultaEx(StrSql,cn,1,parametros)


  if NOT rsvariables.eof   then




   if opcion="FIN" THEN

     sfechai = clng(date()) - clng(cdate("01-01-1900"))
     strsql ="update orvariables set fechact=" & sfechai & " where objid=0" & varz
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     strsql ="update orvariables set fechact=" & sfechai & " where idindicador=0" & id & " and fechact is null"
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)


     strsql ="update orvariables set activa=1 where idindicador=0" & id
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)
     strsql ="update orindicadores  set FECHAULTACT=" & sfechai &" where objid=0" & id
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)
     response.redirect  "oraCTvariables.asp"

   end if


  'Encabezados
  strsql =" select a.* ,b.descri AS DEFF ,b.definicion from orindicadores a,oRformulas b where a.objid=0" & id

  strsql= strsql & " and a.formula=b.objid"
 ' response.write strsql
  set rsind = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rsind.eof then
   UNIDADV=rsind(5)
   UNIDADRES=rsind(11)
   Strsql="select descri from ormedidas where objid=" & UNIDADV
    set rsun = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    UNIDADV=rsun(0)
    call ofv.cerrarconsulta(rsun)

    Strsql="select descri from ormedidas where objid=" & UNIDADRES
	    set rsun = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	    UNIDADRES=rsun(0)
    call ofv.cerrarconsulta(rsun)


   defFormula=rsind("definicion")

   Celda =  ofv.gencelda("","UTi COLSPAN=3 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","",RSIND("DESCRIPCION"),"si")
   Celda = CELDA & ofv.gencelda("","UTi COLSPAN=3 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","","FUNCION :      " & RSIND("DEFF"),"si")
   filaZ= ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda 	 ,"si")
   CELDA=""
  end if
  CALL OFV.CERRARCONSULTA("RSIND")
 sfechai = clng(date()) - clng(cdate("01-01-1900"))
 select case defFormula
 	  case "SUMAR"


         strsql="update  csorvariables2 SET valor=(SELECT SUM(ISNULL(VALOR,0)) as valor  FROM csorvariables2 where varnombre <>'Z' and idindicador=0"& ID & ") WHERE  IDINDICADOR=0"& ID & " AND VARNOMBRE ='Z' "

 	     CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
 	   CASE "PORCENTAJE"
 	      strsql="select * from csorvariables2 WHERE IDINDICADOR=0"&ID& " AND VARNOMBRE <>'Z'"
 	     ' RESPONSE.WRITE STRSQL
 	      set rsfx=ofv.crearconsultaEx(StrSql,cn,1,parametros)
 	      posi=0
 	      do while not rsfx.eof
 	      ArrPorcentaje(POSI)=cdbl("0" & rsfx("valor"))
 	     ' response.write cdbl("0" & rsfx("valor")) & "<br>"
 	      posi=posi +1
 	      rsfx.movenext
 	      loop
          on error resume next
 	      if ArrPorcentaje(1)=0 then ArrPorcentaje(1)=1
 	      result=round((ArrPorcentaje(0)/ ArrPorcentaje(1)) * 100 ,5)

 	      result=replace(result,",",".")

 	      strsql="update  csorvariables2  SET valor='"&result&"' WHERE IDINDICADOR=0"&ID& " AND VARNOMBRE ='Z' "
 	    '  response.write strsql
 	      set rsfx=ofv.crearconsultaEx(StrSql,cn,1,parametros)
 	  CASE "UNICO"

 	     strsql="update  csorvariables2 SET valor=(SELECT ISNULL(VALOR,0)as valor  FROM csorvariables2 where varnombre <>'Z' and idindicador=0"& ID & ") WHERE  IDINDICADOR=0"& ID & " AND VARNOMBRE ='Z' "

 	    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    END SELECT


  Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","","","si")
  Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","NOMBRE","si")
  Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","VALOR","si")
  Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","MEDIDA","si")
  Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","DESCRIPCION","si")
  Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15%","","","","FECHA ACTUALIZACION","si")

  Formulario = celda ' ofv.Encabezados("NOMBRE","VALOR","DESCRIPCION","FECHA&nbsp;ACTUALIZACION")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



  WKSETTABLA = "ORVARIABLES"


  Strsql = "Select * From CSORVARIABLES2  WHERE  IDINDICADOR=0" & id
  'RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  CELDA=""
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi


    IF RS("VARNOMBRE")="Z" THEN
    Celda=""
    Varz=rs(0)
'   if not isnull(rs("fechact")) then
    sInput = "RESULTADO:"
    Celda = ofv.gencelda("","UTi COLSPAN=2 style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","1%","","","",sInput,"si")

   ' Celda = CELDA & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","","si")

    if defFormula="PORCENTAJE" THEN
    sInput = RS("VALOR") & "%"
    ELSE
    sInput = replace(RS("VALOR")&"" , "," ,".")
    END IF

    Celda = CELDA&ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#BBBBBB' ","","","","","",sInput,"si")

     SInput = UNIDADRES
     Celda = CELDA&ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#BBBBBB' ","","","","","",sInput,"si")


    sInput = GTFechadorSH(RS("FEChact"))


     Celda = CELDA & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","","si")
     Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput  ,"si")
'  end if


    ELSE


    sInput = rs(5) & "&nbsp;"
    Celda = ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")



    sInput =RS(9)& "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

     sInput = ofv.GenerarInput(rs(6).name,replace(rs(6) &"",",","."),"text",sForm,"30","dtn","")
     Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    sInput = UNIDADV
    Celda = CELDA&ofv.gencelda("","UTi  ALIGN=RIGHT style='border-style:SOLID;border-width:1;border-color:#BBBBBB' ","","","","","",sInput,"si")

     SInput =RS(4)& "&nbsp;"
     Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


     sInput = GTFechadorSH(RS("FEChact"))
     sfechai = clng(date()) - clng(cdate("01-01-1900"))

	 sInputf = ofv.GenerarInput("fechact",sfechai,"hidden	",sForm,"70","dt","")






     Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput & sinputf,"si")


     END IF


     Fila  =  ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda 	 ,"si")



    xFila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop

      sAccion = ""
      if bModificar then
        sAccion = "orcargavariables.asp?opcion=FIN&varz=" & Varz
      end if
      FILA2 = "<INPUT TYPE='BUTTON'  onclick=""disp_confirm('Confirma ???','FACC');"" VALUE='CONFIRMA RESULTADO' CLASS='BT'> "
      Formu2 = ofv.GenForm("FACC","",sAccion,"","post","",Fila2,"si")
      	sAccion = "orcargavariables.asp?opcion=VERINFO"
    	FILA2 = "<INPUT TYPE=SUBMIT VALUE='VER HISTORIAL' CLASS='BT'> "
     	Formu3 = ofv.GenForm("FACC2","",sAccion,"","post","",Fila2,"si")
     	Celda = ofv.gencelda(""," 	","","","","","",Formu2,"si")
     	Celda = CELDA & ofv.gencelda(""," 	","","","","","",Formu3,"si")
     	xFila = ofv.genrow("","","","","","","",Celda,"si")
     	TABLA2= ofv.GenTabla("","","","10% ","","0","","0","0","","", XFILA,"si")
        sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","", FILAZ & Formulario,"si") & TABLA2 & UMBRALES(ID)


  AGREGAR=""
  buscar=""
  'consulta
  'sAccion = "oraCTvariables.asp"
  'Celda = ofv.BotonVolver(sAccion)

  'sHTML =  sHTML '& ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  ELSE

    if opcion="AGREGARVAR" THEN

   	      strsql="select * from csorvariables2 where idindicador=0" &id & " order BY varnombre "

          SET CVAR=ofv.crearconsultaEx(STRSQL,cn,1,parametros)

          strsql="select max(b.grupo) from ORVARIABLES b where b.idindicador=0" & id
              set rsgrupo= ofv.crearconsultaEx(strsql,cn,1,parametros)

              if not rsgrupo.eof then
                   grupo=clng("0" &  rsgrupo(0)) +1
              else
                   grupo=1
            end if
           call ofv.cerrarconsulta(rsgrupo)
          do while not cvar.eof



            strsqlv=" insert into orvariables (VALOR,grupo,activa,idindicador,VARNOMBRE,descri,TIPOCARGA,NOMBRE,MEDIDA)"
		    strsqlv=strsqlv & "  VALUES (0,"& grupo &",0," & id &",'" & CVAR("VARNOMBRE") & "','"&CVAR("DESCRI")&"','"&CVAR("TIPOCARGA")&"','"& CVAR("NOMBRE")& "','" & CVAR("MEDIDA")&"')"

   	        call ofv.crearconsultaEx(strsqlv,cn,1,parametros)

            cvar.movenext
          loop
          response.redirect "orcargavariables.asp"
    end if
       sAccion = ""
   	'if bModificar then
   	   sAccion = "orcargavariables.asp?opcion=AGREGARVAR"
   	'end if
   	FILA2 = "<INPUT TYPE='BUTTON'  onclick=""disp_confirm('Confirma  ???','FACC');"" VALUE='CONFIRMA NUEVAS VARIABLES' CLASS='BT'> "
   	Formu2 = ofv.GenForm("FACC","",sAccion,"","post","",Fila2,"si")
   	sAccion = "orcargavariables.asp?opcion=VERINFO"
   	FILA2 = "<INPUT TYPE=SUBMIT VALUE='VER HISTORIAL' CLASS='BT'> "
   	Formu3 = ofv.GenForm("FACC2","",sAccion,"","post","",Fila2,"si")

    Celda = ofv.gencelda("","UTi  ","","","","","",Formu2,"si")
    Celda = CELDA & ofv.gencelda("","UTi  ","","","","","",Formu3,"si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    TABLA2=ofv.GenTabla("","","","10% ","","0","","0","0","","",  Fila,"si")

   	sHTML = sHTML &  TABLA2





  end if
END SELECT
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



      sHTML = sHTML & "<script type='text/javascript'>"
      sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
      sHTML = sHTML & "  {"

      sHTML = sHTML & " var i; var r=confirm(disp_msg);"
      sHTML = sHTML & "  if (r==true)"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
     ' sHTML = sHTML &   "     alert(fr.elements['VALOR'].value);"
     ' sHTML = sHTML &  "    for ( i=0; i<fr.elements.length; i++){"
     ' sHTML = sHTML &   "     if (fr.elements[i].name=='VALOR')"
     ' sHTML = sHTML &   "     alert(fr.elements[i].name);}"
     ' sHTML = sHTML &   "                }                     }"
      'sHTML = sHTML &   "     alert(fr.elements['VALOR'].value);                                       }"
      sHTML = sHTML & "    fr.submit();"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  else"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "    alert('Se cancelo la accion!');"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  }"




      sHTML = sHTML & "function send_confirm(send_msg,send_href,send_aid)"
      sHTML = sHTML & "  {"
      sHTML = sHTML & "  var r=confirm(send_msg);"
      sHTML = sHTML & "  if (r==true)"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "    alert('ok ');"
      sHTML = sHTML & "    var fr= document.getElementById(send_aid);"
      sHTML = sHTML & "    fr.href = send_href;"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  else"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "    alert('Se cancelo la accion! ');"
    '  sHTML = sHTML & "    return void(0);"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  }"
      sHTML = sHTML & "</script>"
    IF opcion = "VERINFO"  THEN
     IF  bVOLVER <> "NO" THEN
      sAccion = "orCARGAVARIABLES.asp?"
      Celda = ofv.BotonVolver(sAccion)
     END IF
    ELSE
    if opcion <> "VERINFINDI" then

    sAccion = "oraCTvariables.asp"
     Celda = ofv.BotonVolver(sAccion)
    else
    celda=""
    end if
    END IF
  sHTML =  sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
  sHTML = sHTML  & "</body></html>"
end if

Response.Write sHTML


FUNCTION UMBRALES(SCODINDI)
DIM RS
DIM CELDA
DIM SHTML
DIM Formulario
DIM sForm

  IF OPCION<>"VERINFO" THEN
   strSQl=" select valor from CSorvariables2 where varnombre='Z' AND idindicador=" &SCODINDI
  ELSE
   strSQl=" select valor from CSorvariables2 where varnombre='Z' AND idindicador=" &SCODINDI
  END IF

'  response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  RESULT=REPLACE(RS(0)&"",",",".")
  CALL OFV.CERRARCONSULTA(RS)
  WKSETTABLA = "ORUMBRALES"
  Strsql = "Select A.VALOR1,A.VALOR2  From ORUMBRALES A  WHERE ORDEN IN ((SELECT MIN(B.ORDEN) FROM ORUMBRALES B WHERE B.IDINDICADOR=A.IDINDICADOR),(SELECT MAX(B.ORDEN) FROM ORUMBRALES B WHERE B.IDINDICADOR=A.IDINDICADOR))"
  Strsql = STRSQL & " AND IDINDICADOR="&SCODINDI
  set rsLIMITES = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  IF NOT rsLIMITES.EOF THEN
   MIN=rsLIMITES(0)
   rsLIMITES.MOVENEXT
   MAX=rsLIMITES(1)
  END IF

  CALL OFV.CERRARCONSULTA(rsLIMITES)
  Strsql = "Select * From orumbrales  WHERE VALOR1 <=0"& RESULT & " AND VALOR2>=0"& RESULT & " AND IDINDICADOR="&SCODINDI
  'RESPONSE.WRITE STRSQL

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posI = 0
  Do While (Not rs.EOF)
    sForm = "Uform" & posi




    sInput = RS(1)
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")


    sInput =RS(3) & "&nbsp;&nbsp&nbsp;|&nbsp;&nbsp&nbsp;DESDE:" & rs("valor1") & "&nbsp;&nbsp&nbsp;HASTA:" & rs("valor2")
    Celda = Celda & ofv.gencelda("","ut STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;'","","20%","","","",sInput,"si")

     nCol = 40


           wdescri=""
	       if  not  isnull(rs(6)) then
	           wdescri = TRIM("" &rs(6))
	       end if
'
	       if  len(wdescri) = 0 then wdescri = " "
'
	       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' "
	       t3b = t3b & "name=" & rs(6).name & " cols=" & nCol & " rows=2 "
	       if  bModificar and modifica then
	           t3b = t3b & " onblur=" & sForm & ".submit() "
	           t3b = t3b & " onkeypress=" & chr(34)
	           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
	           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
	       else
	           t3b = t3b & " readonly >"
	       end if
	       xdatos =  wdescri
	       t3b    = t3b & xdatos & "</textarea></font>"
	       SINPUT = t3b



       sInput = "&nbsp;"
       Celda = Celda & ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:" & RS(7) & ";' ","CENTER","5%","","","",sInput,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
    exit do
  Loop
  'sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

    Celda = "<BR>" & ofv.gencelda("","VTI  COLSPAN=9 STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big>UMBRALES</BIG>","si")
    FILAC=ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
    Celda = ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ","","1%","","","","","si")

    Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","	","","","","<big>DESCRIPCION</BIG>","si")

    CELDA = CELDA &  ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big>COLOR</BIG>","si")
    FilaC  = FILAC & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
    if posi=0 then

	 formulario = ofv.gencelda("","UT colspan=3 STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:#ff0000;' ","CENTER","5%","","","","<font color=#ffffff>FUERA DE RANGOS<BR>MINIMO:"& min & "&nbsp;&nbsp;MAXIMO:"&max& "</font>","si")
	 formulario = ofv.genrow("","","","","","","",formulario,"si")
    end if
   UMBRALES = ofv.GenTabla("","","","30% ","","0","","0","0","","",FILAC & Formulario,"si")

END FUNCTION

%>