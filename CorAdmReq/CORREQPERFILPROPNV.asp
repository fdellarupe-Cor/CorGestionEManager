<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 12
  CantRegMoverx = request.querystring("CantRegMover")


  wperfil = request.querystring("perfil")


  IF wperfil <> "" THEN 
     SESSION("wperfil") = wperfil
  ELSE
     wperfil = SESSION("wperfil")
  END IF


  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_perfil=" & Wperfil
  Volver2 = replace(vuelta,"../","")

  sHTML = ofv.FormHeader(session("FormN"))

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "resizable=yes,menubar=no,toolbar=no,height=350,width=350" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"


  'Encabezados
  Formulario = ""

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>L</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=4>&#064;</FONT></B>"
    sImagenCL = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>&#0232;</FONT></B>"
    

  dim stabla(12)


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  sopcion = request.querystring("opcion")

  if sopcion = "SI" THEN
     sECTOR = request.querystring("SECTOR")
     GENSECTORDEST(sECTOR)
  END IF




  posi = 0

  Strsql = "Select * From TIPOREQNV where ID = " & wperfil
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     sform = "reqnv"

     fila = ""


     TREQCOD = RS(1) & ""
     TWKFCOD = RS(5) & ""


      Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","20%","","","","PERFIL: ","si")

      sInput = rs(1) & "&nbsp;-&nbsp;" & rs(2)
      Celda = celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si") & ofv.GenCelda("","uti","","80%","","","","&nbsp;","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



      Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","ACTIVO: ","si")

         scomK = "N" & chr(9) & "S"
         scomD = "NO" & chr(9) & "SI"

         if bModificar then
             sInput = ofv.gencomboFX(sForm,rs(14).name,rs(14),scomK,scomD," ")
         else
             sInput = ofv.GenerarInput("",rs(14),"readonly","","70","dt","")
         end if
      Celda = celda & ofv.GenCelda("","uti","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



      
      ' PERMISO
              Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","20%","","","","RESTRINGIDO: ","si")
      
             scomK = "N" & chr(9) & "S"
             scomD = "NO" & chr(9) & "SI"
            
             if bModificar then
                if rs(13)="S" then
                 sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
		 sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
		 sinput2 = sinput2  & "none" & chr(34) & ";'  href=Corlstusr.asp?PERFIL=" &wperfil&"&cod=" &rs(1)&"&tabla=GPREQXUSR&volver=" & Volver2 
                 sinput2 = sinput2  & " TITLE='Lista Usuarios'  ><b>" & sImagenCl & "</b></a>" 
                 sinput2=  ofv.GenCelda("","uti","","80%","","","",sinput2,"si")
                else
                 sinput2= ofv.GenCelda("","uti","","80%","","","","&nbsp;","si")
                end if
      	        sInput = ofv.gencomboFX(sForm,rs(13).name,rs(13),scomK,scomD," ") & sinput2
             else
      	         sInput = ofv.GenerarInput("",rs(13),"readonly","","70","dt","")
             end if
           Celda = celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
            
           Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
         

      Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","PROYECTO: ","si")

         scomK = "N" & chr(9) & "S"
         scomD = "NO" & chr(9) & "SI"

         if bModificar then
             sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3),scomK,scomD," ")
         else
             sInput = ofv.GenerarInput("",rs(3),"readonly","","70","dt","")
         end if
      Celda = celda & ofv.GenCelda("","uti","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","CUENTA: ","si")

    if bModificar then
         StrSql = "Select * From PROYCUENTAS WHERE ESTADO = 1 ORDER BY ORDEN"
         sInput = ofv.GenerarCombo(sForm,rs(4).NAME,rs(4),strsql,cn,0,1,"")
      else
         StrSql = "Select DESCRIPCION From PROYCUENTAS where ID = " & RS(4)
         set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if NOT rsX2.eof then
            sInput  = rsX2(0)

         ELSE
            sInput  = "SIN ASIGNAR"

         END IF
      end if
      Celda = celda & ofv.GenCelda("","uti","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","CIRCUITO: ","si")

      SEGWHERE = " WHERE VERSION = 2 AND INICIO = 'A' AND ( (OBJTIPO = 'GEN'  )  OR (OBJTIPO = 'REQ')) "

    if bModificar then
        StrSql = "Select * From WKFCIRCUITO " & SEGWHERE & " ORDER BY DESCRIPCION "
        sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5),strsql,cn,1,2,"sb0")
    else
        StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & rs(5) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(2)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

      Celda = celda & ofv.GenCelda("","uti","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



      Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","NIVEL REQUERIDO: ","si")

    if bModificar then
      strsql = "Select * From NivelConfiden Order By NCcod"
      sInput = ofv.GenerarCombo(sForm,rs(10).name,rs(10),strsql,cn2,0,1,"")
    else
      strsql = "Select * From NivelConfiden where NCcod = " & rs(10)
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

      Celda = celda & ofv.GenCelda("","uti","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




      Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","DESTINATARIO: ","si")

                 set cn3 = ofv.conectar(ofv.strconn2)

                 XRESPO = "0" & RS(11)

                 if bModificar then
                    strsql = "Select * From ECO ORDER BY DESCRI"
                    sInput = ofv.GenerarCombo(sForm,RS(11).name,RS(11),strsql,cn3,0,1,"")
                 else
                    strsql = "Select * From ECO where ECO = " & XRESPO
                    set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
                    if not rsniv.eof then sniv = rsniv(1)
                    call ofv.cerrarconsulta(rsniv)
                    sInput = sniv
                 end if

                 call ofv.cerrarconn(cn3)

      Celda = celda & ofv.GenCelda("","uti","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



      Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","NIVEL DESTINATARIO: ","si")

    if bModificar then
      strsql = "Select * From NivelConfiden Order By NCcod"
      sInput = ofv.GenerarCombo(sForm,rs(12).name,rs(12),strsql,cn2,0,1,"")
    else
      strsql = "Select * From NivelConfiden where NCcod = " & rs(12)
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

      Celda = celda & ofv.GenCelda("","uti","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")






           sAccion = ""
           if bModificar then
              sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
              sAccion = sAccion & "&tabla=TIPOREQNV&volver=" & vuelta
           end if
           Formulario = ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

      sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","1","","",Formulario,"si")

'''''''''''' ANALIZA LOS NIVELES EN FUNCION DEL CIRCUITO ASIGNADO

  IF TWKFCOD <> "" THEN
  

     StrsqlT = "Select * From TIPOREQNIVNV where TIPOREQCOD = '" & TREQCOD & "' ORDER BY ITEMCOD "
     set rs1 = ofv.crearconsultaEx(StrSqlT,cn,1,parametros)
     if not rs1.eof then


        IF RS1(3) <> TWKFCOD THEN

           Strsql = "DELETE From TIPOREQNIVNV where TIPOREQCOD = '" & TREQCOD & "'"
           CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

           GENTREQNIVELES()

           set rs1 = ofv.crearconsultaEx(StrSqlT,cn,1,parametros)

        END IF






     ELSE

        GENTREQNIVELES()

        set rs1 = ofv.crearconsultaEx(StrSqlT,cn,1,parametros)


     end if



     if not rs1.eof then




       'Encabezados
        celdas =  ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;'  colspan=8","center","","","","","NIVELES DE RESPONSABLES","si")
        formulario = ofv.GenRow("","","","","22","","",celdas,"si")


        celdas =  ""

        celdas     =  celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","#","si")
        celdas     =  celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","DESCRIPCION","si")
        celdas     =  celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","ESTILO","si")
        celdas     =  celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","MODO","si")
        celdas     =  celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","JERARQUIA","si")
        celdas     =  celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","ESCALA","si")
        celdas     =  celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","TIPO","si")
        celdas     =  celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","CODIGO","si")

        formulario = formulario & ofv.GenRow("","","","","22","","",celdas,"si")




        xTrColor = "#FFFFFF"


        posi = 0
        Do While (Not rs1.EOF) 

     sform = "TREQ" & rs1(0)

     sc = "UT"

     sinput = RS1(4) & ""
     celdas = ofv.GenCelda("","vti  bgcolor=#F7AA22 ","","","","","",sinput,"si")


     if bModificar then
        sInput = ofv.GenerarInput(rs1(5).name,rs1(5),"text",sform,"60","dt","")
     else
        sInput = ofv.GenerarInput("",rs1(5),"readonly","","60","dt","")
     end if
     celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    select case rs1(6) & ""

           case "U"

                 sInput = "UNITARIO"

           case "L"

                 sInput = "LISTA"

           case "LD"

                 sInput = "LISTA Y USUARIOS DEFAULT"

           case else 
           
                 sInput = "sin definir"


    end select 


    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")



        sinputK = ""
        sinputt = ""

    scomK = "F" & chr(9) & "A" & chr(9) & "M"
    scomD = "PREFIJADO" & chr(9) & "ASIGNAR POR APL ORIGEN" & chr(9) & "MANUAL ASIGNADO POR EL RESPONSABLE"

    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs1(7).name,rs1(7),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs1(7),"readonly","","70","dt","")
    end if


        sInput = "ASIGNAR POR APL ORIGEN"

    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


''    sc = "Ut  bgcolor=#EEEEEE "

    set cn2 = ofv.conectar(ofv.strconn0)

    if bModificar then


       SCOMSEP = ""
       scomK   = ""
       scomD   = ""

       strsql = "Select * From NivelConfiden Order By NCcod"
       set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
 
       DO WHILE NOT RSNIV.EOF
           scomK = scomK & SCOMSEP & rsniv(0)
           scomD = scomD & SCOMSEP & rsniv(1)
           SCOMSEP = chr(9)
       RSNIV.MOVENEXT
       LOOP

           scomK = scomK & SCOMSEP & "D"
           scomD = scomD & SCOMSEP & "DEFAULT"



       sInput = ofv.gencomboFX(sForm,RS1(8).name,RS1(8),scomK,scomD,"sb0")


    else
      strsql = "Select * From NivelConfiden where NCcod = " & RS1(8)
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sniv
    end if

    call ofv.cerrarconn(cn2)

 '''   sinput = RS1(8)

       sInput = "DEFAULT"
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

         scomK = "N" & chr(9) & "S"
         scomD = "NO" & chr(9) & "SI"

         if bModificar then
             sInput = ofv.gencomboFX(sForm,RS1(9).name,RS1(9),scomK,scomD," ")
         else
             sInput = ofv.GenerarInput("",RS1(9),"readonly","","10","dt","")
         end if

''    sinput = RS1(9)

       sInput = "SI"
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


         scomK = "USR" & chr(9) & "SEC" & chr(9) & "GRP" & chr(9) & "SOL" & chr(9) & "DES" & chr(9) & ""
         scomD = "USUARIO" & chr(9) & "AREA/SECTOR" & chr(9) & "GRUPO" & chr(9) & "SOLICITANTE" & chr(9) & "DESTINATARIO" & chr(9) & "A DESIGNAR"

         if bModificar then
             sInput = ofv.gencomboFX(sForm,RS1(10).name,RS1(10),scomK,scomD," ")
         else
             sInput = ofv.GenerarInput("",RS1(10),"readonly","","10","dt","")
         end if



 ''''   sinput = RS1(10)

    IF (RS1(10) & "") = "SOL" THEN
       sInput = "SOLICITANTE"
    ELSEIF (RS1(10) & "") = "SOL" THEN
       sInput = "DESTINATARIO"
    END IF


    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


    XRESPO = RS1(11) & ""

    select case rs1(10) & ""

           case "USR"


                 set cn2 = ofv.conectar(ofv.strconn0)


                 if bModificar then
                    strsql = "Select * From USUARIOS"
                    sInput = ofv.GenerarCombo(sForm,RS1(11).name,RS1(11),strsql,cn2,0,1,"")
                 else
                    strsql = "Select * From USUARIOS where ID = '" & XRESPO & "' "
                    set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                    if not rsniv.eof then sniv = rsniv(1)
                    call ofv.cerrarconsulta(rsniv)
                    sInput = sniv
                 end if

                 call ofv.cerrarconn(cn2)


           case "SEC"

                 set cn3 = ofv.conectar(ofv.strconn2)

                 XRESPO = "0" & XRESPO

                 if bModificar then
                    strsql = "Select * From ECO"
                    sInput = ofv.GenerarCombo(sForm,RS1(11).name,RS1(11),strsql,cn3,0,1,"")
                 else
                    strsql = "Select * From ECO where ECO = " & XRESPO
                    set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
                    if not rsniv.eof then sniv = rsniv(1)
                    call ofv.cerrarconsulta(rsniv)
                    sInput = sniv
                 end if

                 call ofv.cerrarconn(cn3)



           case "GRP"


                 set cn2 = ofv.conectar(ofv.strconn0)


                 if bModificar then
                    strsql = "Select * From GRUPOS"
                    sInput = ofv.GenerarCombo(sForm,RS1(11).name,RS1(11),strsql,cn2,0,1,"")
                 else
                    strsql = "Select * From GRUPOS where ID = '" & XRESPO & "' "
                    set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                    if not rsniv.eof then sniv = rsniv(1)
                    call ofv.cerrarconsulta(rsniv)
                    sInput = sniv
                 end if

                 call ofv.cerrarconn(cn2)



           case else 
           
                 sInput = "------"


    end select 





 '''   sinput = RS1(11)
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    filas = ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","25","","",celdas,"si")


       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0" 
       ELSE
          xTrColor = "#FFFFFF" 
       END IF



    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs1(0)," ","*") & "&volver=" & vuelta & "&Tabla=TIPOREQNIVNV"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if


    posi = posi + 1
    rs1.Movenext
  loop






  sHTML = sHTML  & "<center><BR>" & ofv.GenTabla("","aaa style='border-color:#993333;border-style:solid;border-width:1;' ","","98%","","0","","1","1","","",formulario,"si")  & "</center>"

     END IF 




  END IF


  end if



  sAccion = "CORREQPERFIL.asp"
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",Celda)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

FUNCTION   GENTREQNIVELES()


     StrSql = "select * from WKFPSCRESPNV WHERE WKFCOD = '" & TWKFCOD & "' ORDER BY NIVELRSP"
     set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     Do While (Not rsX1.EOF) 


             TREQNIVCOD = GTWKFPSC() & "N" & RSX1(4)


             strsql = "insert into TIPOREQNIVNV ("
             strsql = strsql & "TIPOREQCOD,TIPOREQNIVCOD,WKFCOD,ITEMCOD,DESCRIPCION,"
             strsql = strsql & "MODORSP,ESTILORSP,SUPNIVELRSP,AUTOESCALARSP,"
             strsql = strsql & "TIPORSP,CODIGORSP"
             strsql = strsql & ") values ("
             strsql = strsql & "'" & TREQCOD & "','" & TREQNIVCOD & "','" & TWKFCOD & "','" & RSX1(4) & "','" & RSX1(3) & "',"
             strsql = strsql & "'" & RSX1(5) & "','" & RSX1(6) & "','D','" & RSX1(8) & "',"
             strsql = strsql & "'" & RSX1(9) & "','" & RSX1(10) & "'"
             strsql = strsql & ")"
 '   response.write strsql & "<br>" & chr(13)
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)




       rsX1.Movenext
    loop

END FUNCTION


%>