<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  WSPSESION="and"

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  sLOTE  = request.querystring("LOTE") & ""
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) &"_LOTE=" &  sLOTE
  sopcion  = request.querystring("opcion") & ""
  PORIGEN=REQUEST.QUERYSTRING("PORIGEN")

    sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
    sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>q</FONT></B>"
    sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
    sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
    sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
    sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"

  IF PORIGEN = "" THEN
     PORIGEN = SESSION("PORIGEN")
  ELSE
     SESSION("PORIGEN") = PORIGEN
  END IF


  IF SLOTE = "" THEN
       SLOTE = SESSION("ORLOTE")
    ELSE
       SESSION("ORLOTE") = SLOTE
  END IF




  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ""
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)




 if sopcion = "LOTENEW" then

     sfechai = clng(date()) - clng(cdate("01-01-1900"))

        strsql = "Insert into ORLOTEING "
        strsql = strsql & " (CODIGO,FECHACREATE,FECHAPROCESO,FECHAREVER,ESTADO"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & SESSION.SESSIONID & "','" &  sfechai & "','0','0','1'"
        strsql = strsql & ")"
        STRSQL="PRC_EXECSQL '" & REPLACE(STRSQL,"'","''") & "'"
       ' RESPONSE.WRITE STRSQL & "<br>"
        SET VV= ofv.crearconsultaEx(StrSql,cn,1,volver)
        IDINSERT=VV(0)
        CALL OFV.CERRARCONSULTA(VV)
        STRSQL="UPDATE ORLOTEING SET CODIGO=" & IDINSERT & " WHERE OBJID=" &  IDINSERT
        CALL ofv.crearconsultaEx(StrSql,cn,1,volver)
        STRSQL="UPDATE ORLOTEDET SET LOTEREL=LOTE , LOTE=" & IDINSERT & ",ESTADO=0  WHERE ESTADO=5 " 'AND LOTE=" &  SLOTE
        'RESPONSE.WRITE STRSQL
        CALL ofv.crearconsultaEx(StrSql,cn,1,volver)

  end if


  if sopcion = "LOTEELIM" then

          STRSQL="UPDATE ORLOTEDET SET ESTADO=4  WHERE ESTADO=5 AND LOTE=" &  SLOTE
         ' RESPONSE.WRITE STRSQL
          CALL ofv.crearconsultaEx(StrSql,cn,1,volver)

  end if




'''''''''''''''''''''''''''''''''''''''
'' INCORPORO LOTE NUEVO

  if sopcion = "SI" then

     sfechai = clng(date()) - clng(cdate("01-01-1900"))

        strsql = "Insert into ORLOTEDET "
        strsql = strsql & " ("
        strsql = strsql & "LOTE,FECCONT,FECVALOR,CUENTA,SUCURSAL,"
        strsql = strsql & "MONTO,TEXTO,CLIENTE,FECPROC,ESTADO,"
        strsql = strsql & "INFTIPO"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & SLOTE & "','0','0','','',"
        strsql = strsql &  "'0','DESCRIPCION DEL EVENTO','','0','0','1'"
        strsql = strsql & ")"
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        strsql = "UPDATE ORLOTEING SET ESTADO = '1' WHERE CODIGO = '" & SLOTE & "' AND ESTADO = '0' "
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if




  sinput2="<input type='button' value='ELIMINAR' onclick=""disp_confirm('Confirma eliminar la seleccion???','LOTEELIM');""  CLASS='BT'>"
  sAccion2 ="orinterfasesdetex.asp?opcion=LOTEELIM"
  Formulario2= ofv.genform("LOTEELIM","",sAccion2,"","post ","", sinput2,"si")



  sinput2 = "<INPUT TYPE='button' VALUE='PENDIENTES' onclick=""disp_confirm('Confirma crear Nuevo Lote???','LOTENUEVO');""  CLASS='BT'>"
  sAccion2 ="orinterfasesdetex.asp?opcion=LOTENEW"
  Formulario1 = ofv.GenForm("LOTENUEVO","",sAccion2,"","post","",sinput2,"si")

  celdaz=ofv.gencelda("","ut  xSTYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","",Formulario1,"si")
  celdaz=celdaz & ofv.gencelda("","ut  xSTYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","",Formulario2,"si")
  filax=ofv.genrow("","AAA  ","","","","","",Celdaz,"si")

  xtabla =  "<br><br>" & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filax,"si")
  'filax=""
  'celda=""

  Formulario =  ofv.Encabezados("CODIGO","CREACION","PROCESO","REVERSION","ESTADO")



  WKSETTABLA = "ORLOTEING"

  Strsql = "Select * From " & WKSETTABLA & " WHERE CODIGO = '" & SLOTE & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

    'CODIGO
    sInput = "LOTE:&nbsp;" & RS(1)
    Celda = ofv.gencelda("","CC COLSPAN=2 ","","5%","","","",sInput,"si")

    'FECCREATE
    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(2))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")

    'FECPROCESO
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(3))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")

    'FECREVER
    if clng(rs(4)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(4))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")



   'ESTADO

    wkestado = rs(5) & ""
    select case wkestado

           case "0"

                 wkestadoD = "<font color='#000000' >NUEVO</font>"

           case "1"

                 wkestadoD = "<font color='#993333' >REGISTRADO</font>"


           case "2"

                 wkestadoD = "<font color='#009900' >PROCESADO</font>"


           case "3"

                 wkestadoD = "<font color='#000099' >REVERTIDO</font>"

           case "4"

                 wkestadoD = "<font color='#ff0000' >ELIMINADO</font>"

           case "5"

                 wkestadoD = "<font color='#ff0000' >SELECCIONADO</font>"


           CASE ELSE


                 wkestadoD = "<font color='#000000' >EN REVISION</font>"


    end select

    sInput = wkestadoD
    Celda = Celda & ofv.gencelda("","UTI","","10%","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

   ' Formulario = Formulario & Fila

   END IF
 ' sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si") & "<BR>"

  sImageELS    = "<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
  sImageELN    = "<font face='corflaticon' size=3 color='#dddddd'>&#0253;</font>"
   fila=""
  Formulario = ""

  WKSETTABLA = "ORLOTEDET"


    IF request.FORM.COUNT > 1 THEN
      SESSION("PORIGEN")=""
	  SORLOTE= request.FORM("ORLOTE")
      SESSION(WSPSESION)=""
	  FINICIO=REQUEST.FORM("FINICIO")
	  FFIN=REQUEST.FORM("FFIN")
	  SORCUENTA1=REQUEST.FORM("ORCUENTA1")
	  SORCUENTA2=REQUEST.FORM("ORCUENTA2")
	  SORCUENTA3=REQUEST.FORM("ORCUENTA3")
      SORESTADO=REQUEST.FORM("ORESTADO")

   end if


   IF SORLOTE="" THEN SORLOTE="TODOS"

   IF FINICIO="" THEN FINICIO="TODOS"
   IF FFIN="" THEN FFIN="TODOS"
   IF SORCUENTA1="" THEN SORCUENTA1="TODOS"
   IF SORCUENTA2="" THEN SORCUENTA2="TODOS"
   IF SORCUENTA3="" THEN SORCUENTA3="TODOS"
   IF SORESTADO=""  THEN SORESTADO="TODOS"



  'WSPSESION = "AND"
  'SESSION(WSPSESION) = ""
  SCONECTOR=""


select case SESSION("PORIGEN")

case "INTERFAZ"
    Strsql = "Select  *  From " & WKSETTABLA &   "  WHERE LOTE=0"& SLOTE &" ORDER BY lote DESC"

CASE ELSE
    Strsql = "Select  *  From " & WKSETTABLA & " B "

   IF SORLOTE <> "TODOS" THEN

  	SORCODIGOW = " (B.LOTE = 0" & SORLOTE & ") "

  	SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORCODIGOW
  	SCONECTOR = " AND "

   END IF

     IF FINICIO <> "TODOS" THEN

    	       if IsDate(FINICIO) then
    	          WWSORFECINIH = CLNG(CDATE(FINICIO)) - CLNG(CDATE("01-01-1900"))

    	       else
    	          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
    	       end if

    	       SORFECINIHW   = " (B.FECVALOR > = " & WWSORFECINIH & ") "

    	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
               SCONECTOR=" AND "

        END IF

        IF FFIN <> "TODOS" THEN

    		       if IsDate(FFIN) then
    		          WWSORFECINIH = CLNG(CDATE(FFIN)) -  CLNG(CDATE("01-01-1900"))

    		       else
    		          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
    		       end if

    		       SORFECINIHW   = " (B.FECVALOR  < = " & WWSORFECINIH & ") "

    		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
    		       SCONECTOR=" AND "

      END IF



       sCUENTAS=""
       DELIM=""
       IF SORCUENTA1 <> "TODOS" THEN
        sCUENTAS="'" & SORCUENTA1 &"'" & DELIM
        DELIM=","
       END IF

       IF SORCUENTA2 <> "TODOS" THEN
  	       sCUENTAS= sCUENTAS &  DELIM & "'" & SORCUENTA2 &"'"
  	       DELIM=","
       END IF
       IF SORCUENTA3 <> "TODOS" THEN
  	       sCUENTAS= sCUENTAS &  DELIM & "'" & SORCUENTA3 &"'"
  	       DELIM=","
       END IF



        IF sCUENTAS <> "" THEN

    		       SORFECINIHW   = " (B.CUENTA  IN (" & sCUENTAS & ")) "

    		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
    		       SCONECTOR=" AND "

        END IF


 IF SORESTADO <> "TODOS" THEN



  		       SORFECINIHW   = " (B.ESTADO   = " & SORESTADO & ") "

  		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
  		       SCONECTOR=" AND "

 END IF

   IF session("and" ) <> "" THEN
     swhere = " where "
   ELSE
    swhere = ""
  END IF

  STRSQL =STRSQL & swhere & SESSION("and")
  STRSQL =STRSQL & " ORDER BY b.lote DESC"
END SELECT

  'RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  ilote=""
  sMuestra="NO"
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    Filax  = ""
    wkestado = rs(10) & ""
    wkTIPO   = rs(15) & ""
    sMODAL   = "UPD"
    sDATO    = "text"


    IF wkTIPO   = "2" OR wkestado <> "0" THEN

       sMODAL   = "CSN"
       sDATO    = "readonly"

    END IF


   '  if ilote <> rs("lote") then
   '  celda=ofv.gencelda("","uu colsapn=9 ","","","","","",genlote(rs("lote")),"si")
  '   filax=ofv.genrow("","","","","25","","",Celda,"si")
   '  ilote=rs("lote")
   '  end if

     celda= ofv.gencelda("","UT  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","SELECCIONAR:&nbsp;","si")
     sFormX="sFormX" & RS(0)
     SINPUT = "<input  type=hidden name=" & rs(10).name &" value=" & rs(10) & ">"

     if rs("estado") =5  then
      tilde=" checked "
     else
      tilde=" "
     end if

     sxform="document." & sFormX & "." &  RS(10).name &".value="
     sinput2="<input  "& tilde &" type=checkbox onclick='if (this.checked) {" & sxform & "5 } else {" & sxform & "0 ;} ;document."&sFormX&".submit();'>"
     sAccion2 = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
     sAccion2 = sAccion2 & "&tabla=" & WKSETTABLA & "&volver=" & vuelta

      IF SESSION("PORIGEN")="INTERFAZ" THEN
        celda=  ofv.gencelda("","UT  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","LOTE:&nbsp;","si")
        Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","<B>" &RS("LOTE") &"</B>","si")
        celda= CELDA & ofv.gencelda("","UT  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","LOTE ANTERIOR:&nbsp;","si")
        Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","<B>" &RS("LOTErel") &"&nbsp;</B>","si")
        celda= CELDA & ofv.gencelda("","UT  COLSPAN=5 xstyle='TEXT-ALIGN:CENTER;background-color:#FF0000;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","&nbsp;","si")
      ELSE

       select case cstr(rs("estado")&"")

       case "0" ,"5" ' nuevo o seleccionado
         formobj= ofv.genform(sFormX,"",sAccion2,"","post ","",SINPUT & sinput2,"si")
         celda=CELDA & ofv.GenCelda("","CC style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","2%","","","",formobj,"si")
		 celda= CELDA & ofv.gencelda("","UT  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","LOTE:&nbsp;","si")
		 Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","<B>" &RS("LOTE") &"</B>","si")
		 celda= CELDA & ofv.gencelda("","UT  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","LOTE ANTERIOR:&nbsp;","si")
         Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","<B>" &RS("LOTErel") &"</B>","si")

       case else
        celda=  ofv.gencelda("","UT  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","LOTE:&nbsp;","si")
		Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","<B>" &RS("LOTE") &"</B>","si")
		celda= CELDA & ofv.gencelda("","UT  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","LOTE ANTERIOR:&nbsp;","si")
		Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","<B>" &RS("LOTEREL") &"&nbsp;</B>","si")
        celda= CELDA & ofv.gencelda("","UT  COLSPAN=5 xstyle='TEXT-ALIGN:CENTER;background-color:#FF0000;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","&nbsp;","si")
       end select
     END IF



     FILAZ=ofv.genrow("","","","","25","","",Celda,"si")
     Celda = ofv.GenCelda("","CC style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","2%","","","",RS(0),"si")
     Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","F.CONTABLE:&nbsp;","si")


    'FECCONT
    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(2))
    end if
    SINPUT = dFechan

    IF sMODAL = "CSN" then
       Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    ELSE

       sInputD = ofv.GenerarInput(rs(2).name,rs(2).value,"HIDDEN",sForm,"1","dt","")
       sInput = "<u>" & sinput & "</u>" & sInputD
       Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha CONTABLE.' " & GTCSETCALR(sForm,rs(2).name,"S","N","","",1,0,0,0,0) & " style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    END IF






    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","F.VALOR:&nbsp;","si")

    'FECVALOR
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(3))
    end if
    SINPUT = dFechan

    IF sMODAL = "CSN" then
       Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    ELSE

       sInputD = ofv.GenerarInput(rs(3).name,rs(3).value,"HIDDEN",sForm,"1","dt","")
       sInput = "<u>" & sinput & "</u>" & sInputD
       Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha VALOR CONTABLE.' " & GTCSETCALR(sForm,rs(3).name,"S","N","","",1,0,0,0,0) & " style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    END IF




    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","CUENTA:&nbsp;","si")
    SINPUT = ofv.GenerarInput(rs(4).name,rs(4),sDATO,sForm,"40","dt","")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","F.PROCESO:&nbsp;","si")

    'FECPROC
    if clng(rs(9)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(9))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=1 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")



    Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")

    Celda = Celda & ofv.gencelda("","UT COLSPAN=2 style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","SUCURSAL:&nbsp;","si")
    SINPUT = ofv.GenerarInput(rs(5).name,rs(5),sDATO,sForm,"10","dt","")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","MONTO:&nbsp;","si")
    SINPUT = ofv.GenerarInput(rs(6).name,rs(6),sDATO,sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","CLIENTE:&nbsp;","si")
    SINPUT = ofv.GenerarInput(rs(8).name,rs(8),sDATO,sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")

   'ESTADO
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","ESTADO:&nbsp;","si")

    select case wkestado

           case "0"

                 wkestadoD = "<font color='#000000' >NUEVO</font>"
                 SMuestra="SI"

           case "1"

                 wkestadoD = "<font color='#993333' >REGISTRADO</font>"


           case "2"

                 wkestadoD = "<font color='#009900' >PROCESADO</font>"


           case "3"

                 wkestadoD = "<font color='#000099' >REVERTIDO</font>"

           case "4"

                 wkestadoD = "<font color='#ff0000' >ELIMINADO</font>"
           case "5"
                 wkestadoD = "<font color='#ff0000' >SELECCIONADO</font>"
                 SMuestra="SI"
           CASE ELSE


                 wkestadoD = "<font color='#000000' >EN REVISION</font>"


    end select

    sInput = wkestadoD
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ofv.gencelda("","UT COLSPAN=2 style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","NUMERO ASIENTO:&nbsp;","si")
	SINPUT = RS("NROLOTE")
	Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")
	Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","ID&nbsp;EMPRESA:&nbsp;","si")
	SINPUT =RS("IDEMPRESA")
	Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")
	Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","USUARIO:&nbsp;","si")
	SINPUT = RS("USUARIO")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","DEB/CRED","si")
	SINPUT = RS("DEBCRED")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

     Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","DESCRIPCION:&nbsp;","si")

      nCol = 120

       if  len(wdescri) = 0 then wdescri = "Sin Descripcion"

       if  not  isnull(rs(7)) then
           wdescri = rs(7)
       end if

       t3b = "<font class=at><textarea class=dt "
       t3b = t3b & "name=" & rs(7).name & " cols=" & nCol & " rows=3 "
       if  bModificar AND sMODAL   = "UPD" then
           t3b = t3b & " onblur=" & sForm & ".submit() "
           t3b = t3b & " onkeypress=" & chr(34)
           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
       else
           t3b = t3b & " readonly >"
       end if
       xdatos =  wdescri
       t3b    = t3b & xdatos & "</textarea></font>"
       SINPUT = t3b & ofv.GenerarInput(rs(10).name,"1","HIDDEN",sForm,"1","dt","")

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

   IF WKTIPO = "2" THEN

    Celda = "" 'ofv.gencelda("","","","1","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UT COLSPAN=2 style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","INFO.INTERFASE:&nbsp;","si")


       SINPUT = RS(11)

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","EVENTO:&nbsp;","si")

       sInputx = "-------" & RS(12)

       Strsql = "Select DESCRI From  ORTIPOCATEGO   "
       Strsql = Strsql & " where OBJID = 0" & RS(12)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx


     ''  SINPUT = RS(12)

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","PRODUCTO:&nbsp;","si")

       sInputx = "-------" & rs(13)

       strsql = "Select descri From TIPOPRODUCTO where OBJID = 0" & rs(13)
      ' RESPONSE.WRITE strsql
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx


    ''   SINPUT = RS(13)

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","CARACTER:&nbsp;","si")

       sInputx = "-------" & RS(14)

       strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & RS(14)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx


   ''    SINPUT = RS(14)

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

   ' PROCESO
    Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")
	Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","PROCESO:&nbsp;","si")

	   sInputx = "-------" & RS(14)

	'   Strsql = "Select DESCRI From  OBJETOS   "
	'   Strsql = Strsql & " where codobjeto = 0" & RS("PROCESO")

     if (session("APLMDLOR") & "") = "BE" then

       Strsql = "Select descripcion From TABPROCESOS01 "
       strsql = strsql & " where objid = 0" & RS("PROCESO")

     else

       Strsql = "Select DESCRI From  OBJETOS   "
       Strsql = Strsql & " where CODOBJETO = 0" & RS("PROCESO")

     end if


	   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	   if not rsniv.eof then sInputx = rsniv(0)
	   call ofv.cerrarconsulta(rsniv)

	   sInput = sInputx


   ''    SINPUT = RS(14)

	Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


' SUBPROCESO
    Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")
	Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","SUBPROCESO:&nbsp;","si")

	   sInputx = "-------" & RS(14)

	'   Strsql = "Select DESCRI From  OBJETOS   "
	'   Strsql = Strsql & " where codobjeto = 0" & RS("SUBPROCESO")

     if (session("APLMDLOR") & "") = "BE" then

       Strsql = "Select descripcion From TABPROCESOS01 "
       strsql = strsql & " where objid = 0" & RS("SUBPROCESO")

     else

       Strsql = "Select DESCRI From  OBJETOS   "
       Strsql = Strsql & " where CODOBJETO = 0" & RS("SUBPROCESO")

     end if


	   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	   if not rsniv.eof then sInputx = rsniv(0)
	   call ofv.cerrarconsulta(rsniv)

	   sInput = sInputx


   ''    SINPUT = RS(14)

	Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


    ' AREA /SECTOR
	    Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")
		Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","AREA/SECTOR:&nbsp;","si")

	   sInputx = "-------" & RS(14)

	   Strsql = "Select DESCRI From  ECO   "
	   xx=0
	   if isnumeric(RS("AREASECTOR")) then
	   xx=RS("AREASECTOR")
	   end if
	  Strsql = Strsql & " where ECO = 0" & xx

	   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	   if not rsniv.eof then sInputx = rsniv(0)
	   call ofv.cerrarconsulta(rsniv)

	   sInput = sInputx


	   ''    SINPUT = RS(14)

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")



    ' AREA SOPORTE
	    Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")
		Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","AREA SOPORTE:&nbsp;","si")

	   sInputx = "-------" & RS(14)

	   Strsql = "Select DESCRI From  ECO   "
	  Strsql = Strsql & " where ECO = 0" & RS("AREASOPORTE")

	   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	   if not rsniv.eof then sInputx = rsniv(0)
	   call ofv.cerrarconsulta(rsniv)

	   sInput = sInputx


	   ''    SINPUT = RS(14)

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

   ' LINEA NEGOCIOS
	Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")
	'Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","AREA SOPORTE:&nbsp;","si")
	Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","LINEA DE NEGOCIOS:&nbsp;","si")

   sInputx = "-------" & RS(14)

	Strsql = "Select DESCRIPCION From  ORLINNEGO2   "
  Strsql = Strsql & " where OBJID = 0" & RS("LINEANEGOCIOS")

   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rsniv.eof then sInputx = rsniv(0)
   call ofv.cerrarconsulta(rsniv)

   sInput = sInputx




   Celda = Celda & ofv.gencelda("","UTi COLSPAN=9 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


   Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")












   END IF




    Celda = ""'ofv.gencelda("","","","1","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UT COLSPAN=10  style='background-color:#DDDDDD;border-TOP-style:SOLID;border-TOP-width:2;border-TOP-color:#000000;' ","","","","","","&nbsp;","si")
    Filax  = Filax & ofv.genrow("","","","","5","","",Celda,"si")

    Fila   = Filax

     sAccion = ""
    if bModificar AND sMODAL   = "UPD" then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if


    Formulario = Formulario & FILAZ & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    posi = posi + 1
    rs.Movenext
  Loop
  IF POSI=0  THEN ' SI NO HAY REGISTROS
    Celda = ""
	Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='#F7EEDD' ","center","","","","","NO EXISTEN REGISTROS QUE CUMPLAN CON LA SELECCION ACTUAL","si")
    Fila  = ofv.GenRow("","","","","","","",Celda,"si")
    sHTML = sHTML & "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")
  ' sHTML = sHTML  &  ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")
  ELSE
    IF SMuestra="SI" AND SESSION("PORIGEN")<> "INTERFAZ" THEN
     sHTML = sHTML & xtabla &  ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")
    ELSE
     sHTML = sHTML &   ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")
    END IF
  END IF

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "ORINTERFASESDET.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""
  IF SESSION("PORIGEN")="INTERFAZ" THEN
   sAccion="ORINTERFAZEX.ASP"
  ELSE
  sAccion = "ORBUSINTERFAZ.asp?" '& REQUEST.QUERYSTRING
  END IF
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCCALENDAR("M")
  sHTML = sHTML & "<script type='text/javascript'>"
      sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
      sHTML = sHTML & "  {"
      sHTML = sHTML & "  var r=confirm(disp_msg);"
      sHTML = sHTML & "  if (r==true)"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
     ' sHTML = sHTML & "    alert(fr.name);"
      sHTML = sHTML & "    fr.submit();"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  else"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "    alert('Se cancelo la accion!');"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  }"

      sHTML = sHTML & "function disp_confirm2(disp_msg,disp_frm)"
	  sHTML = sHTML & "  { "
	  sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
     ' sHTML = sHTML & "    alert(fr.name);"
	  sHTML = sHTML & "    fr.submit();"

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

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML


function genlote(SLOTE)
  dim fila,formulario,rs,celda

  Formulariox1 =  ofv.Encabezados("CODIGO","CREACION","PROCESO","REVERSION","ESTADO")



  WKSETTABLA = "ORLOTEING"

  Strsql = "Select * From " & WKSETTABLA & " WHERE CODIGO = '" & SLOTE & "' "
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

    'CODIGO
    sInput = "LOTE:&nbsp;" & RS(1)
    Celda = ofv.gencelda("","CC COLSPAN=2 ","","5%","","","",sInput,"si")

    'FECCREATE
    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(2))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")

    'FECPROCESO
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(3))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")

    'FECREVER
    if clng(rs(4)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(4))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")



   'ESTADO

    wkestado = rs(5) & ""
    select case wkestado

           case "0"

                 wkestadoD = "<font color='#000000' >NUEVO</font>"

           case "1"

                 wkestadoD = "<font color='#993333' >REGISTRADO</font>"


           case "2"

                 wkestadoD = "<font color='#009900' >PROCESADO</font>"


           case "3"

                 wkestadoD = "<font color='#000099' >REVERTIDO</font>"

           case "4"

                 wkestadoD = "<font color='#ff0000' >ELIMINADO</font>"

           case "5"

                 wkestadoD = "<font color='#ff0000' >SELECCIONADO</font>"


           CASE ELSE


                 wkestadoD = "<font color='#000000' >EN REVISION</font>"


    end select

    sInput = wkestadoD
    Celda = Celda & ofv.gencelda("","UTI","","10%","","","",sInput,"si")

    Filax1 = ofv.genrow("","","","","","","",Celda,"si")

    Formulariox1 =  Filax1

   END IF
  genlote= genlote & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulariox1,"si") & "<BR>"

end function

%>