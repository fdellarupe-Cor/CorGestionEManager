<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  sestado = 1
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar


  vuelta = ""


 ' response.write "modif " & bModificar
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  IF NCODIGO <> "" THEN
     session("NCliente") = NCODIGO
  ELSE
     NCODIGO = session("NCliente")
  END IF

  IF sTipo <> "" THEN
     session("TCliente") = sTipo
  ELSE
     sTipo = session("TCliente")
  END IF

  vuelta = ofv.vueltaasp2("../") & vuelta
  vuelta2 = ofv.vueltaasp2("")
  sestado = 9

  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "SELECT * FROM OBJprty  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then sestado = clng(rs("estado"))
  call ofv.cerrarconsulta(rs)

  if CSTR(sestado) > "2" then
     if session("Nomodificar") = "S" then bModificar = False
     response.redirect "coobjwfldet.asp?cod=SI&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo
' Session("Form2") <> "DESARROLLO" 
  elseif CSTR(sestado) = "2"  and Session("Form2") <> "DESARROLLO"  then
     response.redirect "edobjwflconf.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
 'and not (session("segct") = "S")
  else
  if not (session("segct") = "S") then
     if session("Nomodificar") = "S" then bModificar = False
  end if

  sHTML   = ofv.FormHeader(session("FormN"))

  Strsql = "SELECT * FROM OBJWFL  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  rs.eof then 

      strsql = "Insert into OBJWFL (codobjeto, tipobjeto, ordenest) "
      strsql = strsql & " Values (" & nCodigo & ",'" & sTipo & "','1')"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)

  else
   ''  response.write "v " & rs(6) & "<br>"
      if rs(6)  then response.redirect "edObjWflHst.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo 

  end if

  call ofv.cerrarconsulta(rs)

  xActiva = 0
  xObliga = 0

  Strsql = "SELECT * FROM Defwfl"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then

     xActiva = rs(1)
     xObliga = rs(2)

  end if

  call ofv.cerrarconsulta(rs)

  if not xActiva then xObliga = 0

  if xObliga then

     strsql = "update OBJWFL set esqauto = 1, ordenest = 1, chat = 1  "
     Strsql = Strsql & " Where codobjeto=" & nCodigo
     Strsql = Strsql & " And tipobjeto = '" & stipo & "'"

     call ofv.crearconsultaEx(StrSql,cn,1,volver)

 end if

    'Estado
     dim Aest(5)
     Aest(0) = 0
     Aest(1) = "Desarrollo"
     Aest(2) = "Para Aprobar"
     Aest(3) = "Aprobada"
     Aest(4) = "Eliminada"
     
     Strsql = "SELECT * from estados  "
     set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     do until rsx.eof
        Aest(rsx(0)) = Ucase(rsx(1))
        rsx.movenext
     loop
     call ofv.cerrarconsulta(rsx)

  

     Strsql = "SELECT * FROM OBJWFL  "
     Strsql = Strsql & " Where codobjeto=" & nCodigo
     Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    Celda = ofv.gencelda("","ttt colspan=4 aaa ","center","100%","","","","<BIG>CIRCUITO DE AUTORIZACION</BIG>","si")
    Fila  = ofv.GenRow("","","","","30","","",Celda,"si")

    Celda  = ofv.gencelda("","tt","","20%","","","","ESTADO&nbsp;ACTUAL:","si")
    sinput = ofv.gencelda("","vti colspan=3","","","","","",Aest(CLNG(sestado)),"si")
    Celda  = celda & sinput
    Fila   = FILA & ofv.GenRow("","","","","25","","",Celda,"si")

  if not rs.eof then

    sForm = "wfl"
    Celda = ofv.gencelda("","tt ","","20%","","","","AUTORIZANTES:","si")

    if bModificar and Cint(sestado) = 1 and not xobliga then
       sInput = ofv.combosino(sForm & "1",rs(3),rs(3).name)
       sInput = ofv.gencelda("","VTi colspan=2","","10%","","","",sInput,"si")
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "1","",sAccion,"","post","",sInput,"si")
    else
       sval = "NO"
       if rs(3) then sval = "SI"
       sinput = ofv.gencelda("","VTi colspan=2","","10%","","","",sval,"si")
    end if

    Celda = Celda & sinput



    if rs(3) and Cint(sestado) = 1  then
       sInput = ofv.GenerarInput("boton1","Lista","submit",sForm & "1a","5","bt","")
       sInput = ofv.gencelda("","ut","center","10%","","","",sInput,"si")
       sAccion = "edobjwfldet.asp?volver=" & vuelta2
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "1a","",sAccion,"","post","",sInput,"si")
    else

       sInput = ofv.gencelda("","ut","","10%","","","","&nbsp;","si")


    end if

    Celda = Celda & sinput
    Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")

''''
'''' solo si esta activa la accion de lista de autorizantes '''''''''''''''''''''
''''


    if rs(3) then

       sval    = "NO"
       svaldsc = "INDISTINTO"
       if rs(4) then
          sval = "SI"
          svaldsc = "ESTRICTO"
       END IF 

    Celda = ofv.gencelda("","tt ","","20%","","","","Orden:","si")

    if bModificar and Cint(sestado) = 1 and not xobliga then
       sInput = ofv.combosino(sForm & "1b",rs(4),rs(4).name)
       sInput = ofv.gencelda("","VTi colspan=2","","10%","","","",sInput,"si")
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "1b","",sAccion,"","post","",sInput,"si")
    else

       sinput = ofv.gencelda("","VTi colspan=2","","10%","","","",sval,"si")
    end if

    Celda = Celda & sinput

    sInput = ofv.gencelda("","TTT","","10%","","","",svaldsc,"si")
    Celda = Celda & sinput
    Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")

    
''''ordenest = 1,


    if bModificar  and Cint(sestado) = 1   then
     strsql = "update OBJWFL set esqauto = 1,  chat = 1  "
     Strsql = Strsql & " Where codobjeto=" & nCodigo
     Strsql = Strsql & " And tipobjeto = '" & stipo & "'"

     call ofv.crearconsultaEx(StrSql,cn,1,volver)
    end if

 



    'Texto del Mensaje
  esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:7pt;background-color:transparent;} "

     XMESAURL   =  left(request.servervariables("SCRIPT_NAME"),instr(RIGHT(request.servervariables("SCRIPT_NAME"),LEN(request.servervariables("SCRIPT_NAME"))-1),"/"))

     XMESAURL   = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL

       Celda = ofv.gencelda("","ttr colspan=4","LEFT","","","","",XMESAURL ,"si")
     '  Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")



       Celda = ofv.gencelda("","ttr colspan=4","LEFT","","","","","MENSAJE:","si")
       Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")

       if bModificar  and Cint(sestado) = 1 then

          wkfmensaje = rs(7) & ""

          if wkfmensaje = "" then
             wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un proyecto"
             wkfmensaje = wkfmensaje & " normativo pendiente de su aprobación."
             wkfmensaje = wkfmensaje & CHR(13) & CHR(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>" 
             wkfmensaje = UCASE(wkfmensaje)   


             strsql = "update OBJWFL set MSGCHAT = '" & wkfmensaje & "'  "
             Strsql = Strsql & " Where codobjeto=" & nCodigo
             Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
 
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
          end if    

          nCol = 80

          t3b = "<font class=at><textarea CLASS=DT "
          t3b = t3b & "name=" & rs(7).name & " cols=" & nCol & " rows=6 "
          if bModificar then
             t3b = t3b & " onblur=" & sForm & "3.submit() "
             t3b = t3b & " onkeypress=" & chr(34)
             t3b = t3b & " if (this.value.length > 2450) {alert('El texto es demasiado"
             t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
          else
             t3b = t3b & " readonly>"
          end if
          t3b = t3b & wkfmensaje & "</textarea></font>"
          sInput = ofv.gencelda("","ut colspan=4","center","","","","",t3b,"si")
          sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
          sInput =  ofv.GenForm(sForm & "3","",sAccion,"","post","",sInput,"si")
       else
          sinput = ofv.gencelda("","ut colspan=4","center","","","","",rs(7),"si")
       end if

       Celda = sinput
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")


   end if

  '  if rs(3) then
       WKFSIG = "S"
  '  ELSE 
  '    WKFSIG = "N"
  '  END IF 

    if rs(3) then
       WKFSIG2 = "S"
    ELSE 
       WKFSIG2 = "N"
    END IF 

    if rs(4) then
       WKFMODAL = "S"
    ELSE 
       WKFMODAL = "N"
    END IF 

 end if


  call ofv.cerrarconsulta(rs)

  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aaa  ","","70% style='border-style:solid;border-width:1;border-color:#999999;' ","","0","0","-2","1","","",fila,"si")
  sHTML = sHTML & "</center>"

  FILA = ""

  Celda = ofv.gencelda("","ttt colspan=4","","","","","","&nbsp;estado del circuito","si")
  Formulario2 = ofv.GenRow("","","","","30","","",Celda,"si")

  if WKFSIG = "S" THEN

     IF sestado = 1 AND (bModificar   or session("Nomodificarce") <> "S") and Session("Segce") = "S"  then

        Strsql = "SELECT * FROM OBJwfldet  "
        Strsql = Strsql & " Where codobjeto=" & nCodigo
        Strsql = Strsql & " And tipobjeto = '" & stipo & "' order by orden,objid"
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF RS.EOF THEN
           WKFSIGOK = "N"
        ELSE
           WKFSIGOK = "S"
        END IF

   '     IF WKFSIGOK = "S" THEN

           Strsql = "SELECT * FROM TIPOBJETO  "
           Strsql = Strsql & " Where "
           Strsql = Strsql & " tipobjeto = '" & stipo & "' "
           set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           IF RS.EOF THEN
              WKFSIGOK = "N2"
           ELSE
              WKFCKT = RS(19)
           END IF

      '  END IF

        IF WKFSIGOK = "S" or (WKFSIGOK = "N" and WKFSIG2 = "N" ) THEN

           Celda = ""
           IF WKFSIGOK = "N" THEN
              sInput2 = "<b>ATENCION!!</b><BR>EN CASO DE INICIAR, EL CIRCUITO CERRARA CON "
              SINPUT2 = SINPUT2 & " <u><b>" & SESSION("SEGUSUARIO") & "</b></u>  COMO UNICO AUTORIZANTE.<BR>"
           ELSE
              sInput2 = "<b>ATENCION!!</b><BR>EL CIRCUITO ESTA EN CONDICIONES DE SER INICIADO POR "
              SINPUT2 = SINPUT2 & " <u><b>" & SESSION("SEGUSUARIO") & "</b></u>  COMO ORIGINANTE.<BR>"
           END IF
           Celda = ofv.gencelda("","uti colspan=3 style='text-align:justify;' ","","","","","",SINPUT2,"si")

           sInput = ofv.GenerarInput("boton1","INICIAR","submit","edPropiedades2","5","bt","")
           Celda = celda & ofv.gencelda("","uti","center","10%","","","",sInput,"si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")

           sAccion = "EDWKFSTART.asp?OPCION=AUTO&MODAL=" & WKFMODAL
           Formulario2 = Formulario2 &  ofv.GenForm("edPropiedades2","",sAccion,"","post","",Fila,"si")

        ELSE 

           Celda = ""
           IF WKFSIGOK = "N" THEN
              sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO, NO INFORMO LA LISTA DE AUTORIZANTES.<BR>"
           ELSE
              sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO, EL TIPO DE DOCUMENTO NO ESTA CONFIGURADO.<BR>"
           END IF
           Celda = ofv.gencelda("","uti colspan=3 style='text-align:justify;' ","","","","","",SINPUT,"si")

           sInput = "&nbsp;"
           Celda = celda & ofv.gencelda("","ut","center","5%","","","",sInput,"si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")

           Formulario2 = Formulario2 &  Fila

        END IF

    ELSE

           Celda = ""

          IF sestado = 1  then
             sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO DESDE AQUI O NO ESTA AUTORIZADO PARA INICIARLO.<BR>"
          else
             sInput = "<FONT COLOR=#005500>EN CURSO.</FONT><BR>"
          end if   
           Celda = ofv.gencelda("","uti colspan=3 style='text-align:justify;' ","","","","","",SINPUT,"si")

           sInput = "&nbsp;"
           Celda = celda & ofv.gencelda("","ut","center","5%","","","",sInput,"si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")

           Formulario2 = Formulario2 &  Fila


    END IF

     IF sestado = 2 AND  Session("Segce") = "S"  then

           sInput = ofv.GenerarInput("boton1","Consultar historial","submit","edPropiedades3","5","bt","")
           sAccion = "CoHistWkf.asp"
           sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
           Celda = ofv.gencelda("","Vt colspan=2","center","30%","","","",sInput,"si")

           if (bModificar   or session("Nomodificarce") <> "S") and Session("Segce") = "S" then
              sInput = ofv.GenerarInput("boton1","Cancelar Circuito","submit","edPropiedades3","5","bt style='background-color:#993333;' ","")
              sAccion = "EDCANCELWkf.asp?OPCION=CONF"
              sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
           else 
              sInput = "&nbsp;"
           end if

           Celda = celda & ofv.gencelda("","Vt colspan=2 ","center","30%","","","",sInput,"si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")


           Formulario2 = Formulario2 & Fila


     end if


  ELSE


     Strsql = "SELECT OBJPRTY.*, Estados.estado FROM OBJPRTY, Estados "
     Strsql = Strsql & "WHERE OBJPRTY.estado = Estados.objid AND codobjeto=" & nCodigo
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        if session("segct") = "S" then
       Celda = ofv.gencelda("","Vt colspan=2","center","30%","","","","CAMBIO DE ESTADO","si")
       Celda = celda & ofv.gencelda("","Vt ","","30%","","","","NUEVO ESTADO:","si")
'xxxxxxxxxxxxx


    Select case cstr(rs(9))
      case "1"
        if (bModificar   or session("Nomodificarce") <> "S") and Session("Segce") = "S"  then
          sInput = "<font class=at><select name=" & rs(9).name
          sInput = sInput & " onchange=document.edPropiedades2.submit(); >"
          sInput = sInput & "<option selected value=1>" & Aest(1) & "</option>"
          sInput = sInput & "<option value=2>" & Aest(2) & "</option></select></font>"
        else
          sInput = ofv.GenerarInput(rs(9).name,rs(16),"readonly","","20","dt","")
        end if
      case "2"
        if (bModificar   or session("Nomodificarce") <> "S") and Session("Segce") = "S" then
          sInput = "<font class=at><select name=" & rs(9).name
          sInput = sInput & " onchange=document.edPropiedades2.submit(); >"
          sInput = sInput & "<option selected value=2>" & Aest(2) & "</option>"
          sInput = sInput & "<option value=1>" & Aest(1) & "</option>"
          sInput = sInput & "</select></font>"
        else
          sInput = ofv.GenerarInput(rs(9).name,rs(16),"readonly","","20","dt","")
        end if
    end select
    Celda = Celda & ofv.GenCelda("","VT","","","","","",sInput,"si")
    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = ""
    if (bModificar   or session("Nomodificarce") <> "S") and Session("Segce") = "S" then
      sAccion = "asp/ActualizarEN.asp?opc=WFL&ID=" & rs(0) & "&Tabla=Objprty&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if
    Formulario2 = ofv.GenForm("edPropiedades2","",sAccion,"","post","",Fila,"si")

   end if
  end if

'xxxxxxxxxx

  END IF



  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#999999;' ","","70% ","","0","0","1","0","","",Formulario2,"si")
  sHTML = sHTML & "</center>"

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Retornar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  sestado = 1
  sHTML = sHTML & "</body></html>"
 end if
end if

  if Cint(sestado) = 1 then response.write sHTML

%>