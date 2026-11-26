<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok 
else
  response.expires=0

 ''''' RESPONSE.WRITE REQUEST.QUERYSTRING
  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))


  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  Volver2 = replace(vuelta,"../","")
  set cn = ofv.conectar(ofv.strconn6)


  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")


 if sopcion = "SI" then
    GENRESP()
 end if



    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>&#0232;</FONT></B>"

    celdas = ""


  'Encabezados
  formulario = ""
  formulario = formulario & ofv.Encabezados("CODIGO","USUARIO","CARGA&nbsp;<BR>&nbsp;RESPONSABLE","CARGA&nbsp;PARA<BR>&nbsp;OTRO","VISTAS<BR>&nbsp;TEMPORALES","VISTAS<BR>&nbsp;NUEVAS","SECTOR","&nbsp;")


  StrSql = "Select * From GPREQRESPONSABLES "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "GPRESP" & rs(0)


       sc = "ut"

    'Eliminar
    if bSuprimir and rs(0) <> 1  then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=GPREQRESPONSABLES&volver=" & vuelta & "&Eliminar=NSNC"
      Celdas = ofv.BotonEliminar(rs(0),sHref)
    else
      Celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

 '   celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")



       sInput = rs(0)
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")



   
       if rs(1)  = "U" then
          sInput = "USUARIO"
       else
          sInput = "SECTOR"
       end if
    '   celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


     
      sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "none" & chr(34) & ";'  href=agregarlista.asp?TABLA=GPREQRESPONSABLES&CodOBJETO=" &rs(0) & "&VOLVER=" & Volver2
      sinput2 = sinput2  & " TITLE='SUBTIPOS DE " & rs(1) & "'  ><b>" & sImagenC & "</b></a>" 
     
    '   celdas = celdas & ofv.GenCelda("",sc,"CENTER","","","","",sInput2,"si")
       if rs(2)<>"TODOS" THEN
        call getInfoUser(rs(2))  
       ELSE
        session("SGU_nombre")="TODOS"
       END IF
       
       if bModificar AND rs(0) <>1 then
       
          sInput = session("SGU_nombre") & "&nbsp;"  & sinput2
       else
          sInput = ofv.GenerarInput("",session("SGU_nombre"),"readonly","","60","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput ,"si")



       scomK = "S" & chr(9) & "N"
       scomD = "SI" & chr(9) & "NO"
       
       
      if  bModificar  then
	 sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3),scomK,scomD," ")

      else
	 if trim(rs(3)&"") = "S" then sInput = "SI"
	 if trim(rs(3)&"") = "N" then sInput = "NO"
      end if
       
      
       
      celdas = celdas & ofv.GenCelda("",sc,"CENTER","5%","","","",sInput,"si")


     
     scomK = "S" & chr(9) & "N"
     scomD = "SI" & chr(9) & "NO"
            
            
     if  bModificar  then
     	 sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD," ")
     
     else
     	 if trim(rs(4)&"") = "S" then sInput = "SI"
     	 if trim(rs(4)&"") = "N" then sInput = "NO"
     end if
            
           
            
      celdas = celdas & ofv.GenCelda("",sc,"CENTER","5%","","","",sInput,"si")

     scomK = "S" & chr(9) & "N"
     scomD = "SI" & chr(9) & "NO"
            
            
     if  bModificar  then
     	 sInput = ofv.gencomboFX(sForm,rs(6).name,rs(6),scomK,scomD," ")
     
     else
     	 if trim(rs(6)&"") = "S" then sInput = "SI"
     	 if trim(rs(6)&"") = "N" then sInput = "NO"
     end if
            
           
            
      celdas = celdas & ofv.GenCelda("",sc,"CENTER","5%","","","",sInput,"si")

     scomK = "S" & chr(9) & "N"
     scomD = "SI" & chr(9) & "NO"
            
            
     if  bModificar  then
     	 sInput = ofv.gencomboFX(sForm,rs(7).name,rs(7),scomK,scomD," ")
     
     else
     	 if trim(rs(7)&"") = "S" then sInput = "SI"
     	 if trim(rs(7)&"") = "N" then sInput = "NO"
     end if
            
           
            
      celdas = celdas & ofv.GenCelda("",sc,"CENTER","5%","","","",sInput,"si")


                set cn3 = ofv.conectar(ofv.strconn2)

                 XRESPO = "0" & RS(8)

                 if bModificar then
                    strsql = "Select * From ECO ORDER BY DESCRI"
                    sInput = ofv.GenerarCombo(sForm,RS(8).name,RS(8),strsql,cn3,0,1,"")
                 else
                    strsql = "Select * From ECO where ECO = " & XRESPO
                    set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
                    if not rsniv.eof then sniv = rsniv(1)
                    call ofv.cerrarconsulta(rsniv)
                    sInput = sniv
                 end if

                 call ofv.cerrarconn(cn3)
            
      celdas = celdas & ofv.GenCelda("",sc,"","5%","","","",sInput,"si")


     sAccion = "GPRESPONSABLESJER.ASP?CODUSR=" & RS(5) & "&CantRegMover=" & CantRegMoverX

     sFUNCION = ""

     wkproccod1 = "<A CLASS=NT " & sFUNCION & " style='COLOR:#ffffff;' TITLE='Nivel Superior.' HREF='" & sAccion & "' >NIV.SUPERIOR</A>"

     wkproccod1S = " style='background-color:#F7AA22;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "

     celdas = celdas & ofv.GenCelda("",sc & wkproccod1S,"CENTER","5%","","","",wkproccod1,"si")



   
    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=GPREQRESPONSABLES"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
    rs.Movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "GPRESPONSABLES.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""
  

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML

FUNCTION GENRESP()
   

    strsql = "insert into GPREQRESPONSABLES ("
    strsql = strsql & "CODUSR,TIPOUSR,PERMITEMR,PERMITEMO,"
    strsql = strsql & "PERMITEVT,PERMITEVN,SECTOR"
    strsql = strsql & ") values ("
    strsql = strsql & "'"& GTrnd("RESP") & "','U','N','N',"
    strsql = strsql & "'N','N','0'"
    strsql = strsql & ")"
    
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION



%>