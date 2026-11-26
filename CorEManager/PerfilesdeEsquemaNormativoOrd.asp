<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = request.querystring("CantRegMostrar")
  CantRegMoverx = request.querystring("CantRegMover")

  wperfil = request.querystring("wperfil")
  welem = request.querystring("welem")

  wOPC = request.querystring("OPC")

''response.write wOPC & " wOPC " & "<br>"

  wwobj = request.querystring("Wobj")

  if len(welem) = 0 then welem = 0

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ""


  dim stabla(15)
  dim stablaE(15)
  dim stablaS(15)

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  posi = 0


  Strsql = "Select * From Objperfil where tipoperfil = '" & wperfil & "' "
 'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     fila = ""


      Celda = ofv.GenCelda("","tt  ","","","","","","Perfil: ","si")
      sInput = rs(0) & "&nbsp;-&nbsp;" & rs(1)

      Celda = celda & ofv.GenCelda("","cc  ","","90%","","","",sInput,"si")

      sinput = "<a  href=PerfilesdeEsquemaNormativoProp.asp?CantRegMover=" & CantRegMoverx
      sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx   & "&Wperfil=" & Wperfil & " >Volver</a>"
      Celda = Celda & ofv.genCelda("","bietx style='display:flex;'","center","","","","",sInput,"si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")

      Celda = ofv.GenCelda("","aaa colspan=3  ","","","","","","&nbsp;","si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      sInput = "ORDEN&nbsp;DE&nbsp;SOLAPAS"
      Celda = ofv.GenCelda("","ttt colspan=2 ","","","","","",sInput,"si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      Formulario = Fila


'     sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")
'     Formulario = ""

''response.write wOPC & "<br>"

'if wOPC= "SEQ" then

 winput = request.form(1) &""

'response.write winput

' Pasa valores a tabla de entrada
' Indice 1 = secuencia de entrada
' Indice 2 = valor en tabla del campo que se muestra
    dim stablaZS(16,1)
    xx = 0
    ii = 1
    yy = 0
    nn = 0
    for xx = 1 to 15

           yy = yy + 1

           hh = xx * 10

           'stablaZS(xx,0) = hh
           'stablaZS(xx,1) = rs(xx + 1 + 30)
           stablaZS(rs(xx + 1 + 30),0) = hh
           stablaZS(rs(xx + 1 + 30),1) = xx
       '   response.write stablaZS(xx,0) & " x=" & stablaZS(xx,1) & "<br>"


    next
    tyy = yy


    '------------------------------------------------------------------------

   if wOPC= "SEQ" then



   xx = 0
   ii = 0
   yy = 0



'Reordenamiento de la tabla de entrada en la salida

   comodins = winput
   comodinv = wwobj
   worka = 0

   worka = stablaZS(wwobj,1)




   wwobjS = wwobj * 10

  ''response.write wwobj & " wwobj " & "<br>"

 ''response.write winput & " winput "& "<br>"

   if cint(wwobjS) < cint(winput) then

      for ii=wwobjS to 150 step 10

          if ii > cint(winput) then
             xx = ii-10
             exit for
          end if
      next

      if xx = 0 then xx = 150

      for ii=wwobjs to xx step 10

	      stablaZS((ii/10),1) = stablaZS((ii/10)+1,1)

	  next

      stablaZS((xx/10),1) = worka

   else

   '' response.write wwobjs & " wwobjs "& "<br>"

     for ii=0 to wwobjs step 10

             if ii > cint(winput) then
                xx = ii
                exit for
             end if
         next

      if xx = 0 then xx = wwobjS

        ''response.write xx & " xx "& "<br>"

      for ii=wwobjs to xx step -10

           stablaZS((ii/10),1) = stablaZS((ii/10)-1,1)

      next

      stablaZS((xx/10),1) = worka

   end if



   tyy = 15

  'response.write tyy & " tyy " & "<br>"
         Strsql = "UPDATE OBJPERFIL SET "

         for xx = 1 to tyy

             Strsql = strsql & rs(stablaZS(xx,1) + 31).name & " = " & xx

             if xx <> tyy then Strsql = strsql & ", "

         next

         Strsql = strsql & " where tipoperfil = '" & wperfil & "' "
       'response.write strsql & " UPDATE " & "<br>"
         call ofv.crearconsultaEx(StrSql,cn,1,parametros)


         Strsql = "Select * From Objperfil where tipoperfil = '" & wperfil & "' "
         ''response.write strsql & " SELECT " & "<br>"
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs.eof then



          for xx = 1 to 15

           'stablaZS(xx,0) = xx
           'stablaZS(xx,1) = rs(xx + 31)
           stablaZS(rs(xx + 31),0) = xx *10
           stablaZS(rs(xx + 31),1) = xx
 ''''response.write stablaS(xx) & " stablaS " & "<br>"
          next

        end if

 '-------------------------------------------
 END IF

     tyy = 15
     xx = 0

     for yy = 1 to tyy
         sform = "formxxx" & yy
         fila = ""

         xx = xx + 1

     '   ssInput = ofv.GenerarInput("S" & yy,yy * 10,"text",sForm,"10","dtn","")
     '   Celda = ofv.GenCelda("","cc  ","right","","","","",ssInput,"si")
     '   sInput = rs(stablaS(yy) + 1 + 15)

        ssInput = ofv.GenerarInput("S" & yy,yy * 10,"text",sForm,"10","dtn","")
		Celda = ofv.GenCelda("","cc  ","right","","","","",ssInput,"si")
        sInput = rs(stablaZS(yy,1) + 1 + 15)
       'response.write yy & " yy en el for " & "<br>"

  if rs(stablaZS(yy,1) + 1) then

      if xx = 1 or xx = 82 then

         Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")
        'response.write xx & " xx por if " & "<br>"
        'response.write sInput & " Sinput por if " & "<br>"
         'Celda = Celda & ofv.gencelda("","cc ","center","","15%","","","","si")

      else

         Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")
        'response.write xx & " xx por else " & "<br>"
        'response.write sInput & " Sinput por else " & "<br>"
         sinput = ""
         'Celda = Celda & ofv.gencelda("","cc","center","","15%","","",sInput,"si")

      end if

  else

      Celda = celda & ofv.GenCelda("","cc ","","60%","","","",sInput,"si")
      'Celda = Celda & ofv.gencelda("","cc","center","","15%","","","","si")

  end if

   ''  Wobj =  stablaS(yy) '''+ 15



     sAccion = ""
	    if bModificar then

	       Wobj =  yy

	       sAccion = "PerfilesdeEsquemaNormativoOrd.asp?CantRegMover=" & CantRegMoverx & "&OPC=SEQ" & "&wobj=" & Wobj
           sAccion = sAccion  & "&CantRegMostrar=" & CantRegMostrarx   & "&Wperfil=" & Wperfil    & "&Welem=" & (stablaZS(yy,1) + 1)

       '' response.write stablaS(yy) & " stablaS(yy+1) " & "<br>"



    end if

	  Fila = ofv.genform(sForm,"",sAccion,"","post","",Celda,"si") & chr(13)
	  Formulario = Formulario & ofv.genrow("","","","","","","",Fila,"si")

      posi = posi + 1


     next

  end if
  sHTML = sHTML & ofv.GenTabla("","aaa","","80% ","","0","","0","2","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

'end if

Response.Write sHTML

%>