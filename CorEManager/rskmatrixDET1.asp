<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  WKRSKCAT = request.querystring("CATEGO") & ""
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

     smenu   = Request.querystring("Menu")
     sorden  = Request.querystring("sorden")
     sNombre = request.querystring("Nombre")


            sID    = Replace(Request.Form("buscar") & "", "'", "")
            sID    = Replace(sID, Chr(34), "")

            sID2   = Request.Form("lista") & ""

            tsel   = Request.Form("tsel") & ""
            tgrp   = Request.Form("tgrp") & ""
            tcod   = Request.Form("tcod") & ""
            ttip   = Request.Form("ttip") & ""
            tnam   = Request.Form("tnam") & ""

            tdoc   = Request.Form("tdoc") & ""


            Addqry = "&buscar=" & replace(sID," ","_")

            Addqry = Addqry & "&Lista=" & sID2

            Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

            Addqry = Addqry & "&tdoc=" & tdoc 
            
            sID    = Replace(sID, "_", " ")





  sHTML = ofv.FormHeader(session("FormN"))




  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From RSKSITUACION WHERE ID = 0" & tdoc
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     SNOMEMPRE = RS(1)

  END IF

  if tcod = "99999" THEN

         SNOMEMPRE2 = "todas las categorias"
       '  sid2 = "1"
  ELSE


       Strsql = "Select * From RSKCATEGO WHERE ID = 0" & tcod
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs.eof then 

          SNOMEMPRE2 = RS(2)

       END IF

  END IF 


  if sid2 = "99999" THEN

         SNOMEMPRE3 = "PONDERACION GLOBAL"
       '  sid2 = "1"
  ELSE

      Strsql = "Select * From RSKitem WHERE ID = 0" & sid2
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rs.eof then 

         SNOMEMPRE3 = RS(1)

      END IF

  END IF 


    sImagen4 = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>4</FONT></B>"
    sImagenP = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"

  Celda = ofv.gencelda("","ttt","","","","","","<BIG>" & sImagen4 & " " & SNOMEMPRE & " / " & SNOMEMPRE2  & " / " & SNOMEMPRE3 & "</BIG>","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","0","0","","",Fila,"si") & "<BR>"



 IF SID2 = "99999" THEN


  'Encabezados
  Formulario = ""

  dim atabitm(200,5)
  toptabitm = 0

    sInput = "#"
    Celda = ofv.GenCelda("","ttr rowspan=2","","2%","","","",posi,"si")

    sInput = "Evento"
    Celda = Celda & ofv.gencelda("","ttr rowspan=2 ","","35%","","","",sInput,"si")


  Strsql = "select  max(itemrsk) as itemrsk, max(itedes) as itedes, max(ponderador) as ponderador  from csmatrizrsk "

  if tcod = "99999" then
     Strsql = Strsql & " where sitid = 0" & tdoc & " "
  else

     Strsql = Strsql & " where catid = 0" & tcod & " "
  end if


  Strsql = Strsql & " group by itedes "
  Strsql = Strsql & " order by ponderador desc, itedes "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 1
  Do While (Not rs.EOF) 

   atabitm(posi,0) = rs(0)
   atabitm(posi,1) = rs(1)
   atabitm(posi,2) = rs(2)
   atabitm(posi,3) = 0
   atabitm(posi,4) = 0
   atabitm(posi,5) = 0
   toptabitm = posi

    sInput = rs(1)
    Celda = Celda & ofv.gencelda("","ttr colspan=2 style='border-left-style:solid;border-left-width:1;border-left-color:#993333;' ","","30%","","","",sInput,"si")


    sInput = "Total"
    Celda2 = Celda2 & ofv.gencelda("","ttr style='border-left-style:solid;border-left-width:1;border-left-color:#993333;' ","","10%","","","",sInput,"si")

    sInput = "Prom"
    Celda2 = Celda2 & ofv.gencelda("","ttr","","10%","","","",sInput,"si")

    sInput = "Max"
  '  Celda2 = Celda2 & ofv.gencelda("","ttr","","10%","","","",sInput,"si")

    posi = posi + 1
  rs.movenext
  loop

    Formulario = Formulario & "<THEADER>"
    Formulario = Formulario & ofv.genrow("","","","","","","",Celda,"si")
    Formulario = Formulario & ofv.genrow("","","","","","","",Celda2,"si")
    Formulario = Formulario & "</THEADER>"
    Formulario = Formulario & "<TBODY>"


  Strsql = "select max(evedes), max(itedes), min(ponderacion) as minimo, avg(ponderacion) as promedio, "
  Strsql = Strsql & "   max(ponderacion) as maximo, max(b.prompond) as prompond , max(a.ponderador) as ponderador, "
  Strsql = Strsql & "   max(a.itemrsk) as coditem, sum(ponderacion) as ponderasum, max(a.catid) as categoria"
  Strsql = Strsql & "  from csmatrizrsk a inner join  "
  Strsql = Strsql & " (select max(evenrsk) as codeven, max(evedes) as evendsc,  sum(ponderacion * ponderador) as prompond "
  Strsql = Strsql & "  from csmatrizrsk "

  if tcod = "99999" then
     Strsql = Strsql & " where sitid = 0" & tdoc & " "
  else

     Strsql = Strsql & " where catid = 0" & tcod & " "
  end if


  Strsql = Strsql & "  group by evedes ) b "
  Strsql = Strsql & "  on a.evenrsk = b.codeven "

  if tcod = "99999" then
     Strsql = Strsql & " where a.sitid = 0" & tdoc & " "
  else

     Strsql = Strsql & " where a.catid = 0" & tcod & " "
  end if


  Strsql = Strsql & " group by a.evedes,a.itedes "
  Strsql = Strsql & "  order by b.prompond desc, a.evedes, a.ponderador desc "




'''response.write Strsql & "<br>"

xTrColor = "#E0E0FF"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  Do While (Not rs.EOF) 
    sForm = "form" & posi

    posi2 = posi + 1


    Celda = ofv.GenCelda("","ut style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","","","","","",posi2,"si")

    sInput = rs(0)

    if tcod = "99999" THEN

       Strsql = "Select * From RSKCATEGO WHERE ID = 0" & rs(9)
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs1.eof then 

           sInput = sInput & "<br><small>(" & rs1(2) & ")</small>"

       END IF
  end if






    Celda = Celda & ofv.gencelda("","ut style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","","","","","",sInput,"si")


    positab = 1
    codevenant = rs(0) 

  ' atabitm(posi,0) = rs(0) cod item
  ' atabitm(posi,1) = rs(1) nom item
  ' atabitm(posi,2) = rs(2) pond item
  ' atabitm(posi,3) = 0 min
  ' atabitm(posi,4) = 0 pro
  ' atabitm(posi,5) = 0 max
  ' toptabitm = posi


    for indtab=1 to toptabitm

        if positab = 1 then 
           coditem = rs(7) 
        else
           coditem = 99999
        end if 
             



        if atabitm(indtab,0) = coditem then

           sInput1 = rs(8) & "&nbsp;"
           sInput2 = rs(3) & "&nbsp;"
           sInput3 = rs(4) & "&nbsp;"

           rs.Movenext
           if rs.eof then
              positab = 2
           else
              if codevenant <> rs(0) then positab = 2
           end if 


        else

           sInput1 = "---&nbsp;"
           sInput2 = "---&nbsp;"
           sInput3 = "---&nbsp;"

        end if


           Celda = Celda & ofv.gencelda("","ut style='border-left-style:solid;border-left-width:1;border-left-color:#993333;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","right","","","","",sInput1,"si")
           Celda = Celda & ofv.gencelda("","vt style='color:#993333;border-left-style:solid;border-left-width:1;border-left-color:#dddddd;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","right","","","","",sInput2,"si")
        '   Celda = Celda & ofv.gencelda("","ut style='border-left-style:solid;border-left-width:1;border-left-color:#dddddd;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","right","","","","",sInput3,"si")




    next



    Formulario = Formulario & ofv.genrow("","aaa bgcolor=" & xTrColor & "  ","","","30","","",Celda,"si")


                         if xTrColor = "#FFFFFF" OR xTrColor = "#E0E0FF" then
                            xTrColor = "#F0F0F0" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF




    posi = posi + 1

  Loop
  sHTML = sHTML & "<center><div  style='background-color:#f9f9f9;overflow:auto; width:900px; height:400px;border-style:solid;border-width:1;border-color:#336699;'>"
    Formulario = Formulario & "</TBODY>"
  sHTML = sHTML & ofv.GenTabla("","","","","","","","-2","2","","",Formulario,"si")
  sHTML = sHTML & "</div></center>"
 ELSE


  'Encabezados
  Formulario = ""

    sInput = "#"
    Celda = ofv.GenCelda("","ttr","","5%","","","",posi,"si")

    sInput = "Evento"
    Celda = Celda & ofv.gencelda("","ttr","","65%","","","",sInput,"si")

    sInput = "Minimo"
    Celda = Celda & ofv.gencelda("","ttr","","10%","","","",sInput,"si")

    sInput = "Promedio"
    Celda = Celda & ofv.gencelda("","ttr","","10%","","","",sInput,"si")

    sInput = "Maximo"
    Celda = Celda & ofv.gencelda("","ttr","","10%","","","",sInput,"si")


    Formulario = Formulario & ofv.genrow("","","","","","","",Celda,"si")


  Strsql = "select max(evedes), max(itedes), min(ponderacion) as minimo, avg(ponderacion) as promedio, "
  Strsql = Strsql & " max(ponderacion) as maximo "
  Strsql = Strsql & "from csmatrizrsk "
  Strsql = Strsql & "where ITEMRSK = 0" & sid2 & " and catid = 0" & tcod & " "
  Strsql = Strsql & "group by evedes,itedes "

  Strsql = Strsql & "order by promedio desc, evedes,itedes "

'''response.write Strsql & "<br>"

xTrColor = "#FFFFFF"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  Do While (Not rs.EOF) 
    sForm = "form" & posi

    posi2 = posi + 1

    Celda = ofv.GenCelda("","ut","","","","","",posi2,"si")

    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")

    sInput = rs(2) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","ut","right","","","","",sInput,"si")

    sInput = rs(3) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","vt style='color:#993333;' ","right","","","","",sInput,"si")

    sInput = rs(4) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","ut","right","","","","",sInput,"si")

    Formulario = Formulario & ofv.genrow("","aaa bgcolor=" & xTrColor & " ","","","30","","",Celda,"si")


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF



    posi = posi + 1
    rs.Movenext
  Loop

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

 END IF




  
 
    sAccion = "rskmatrix.asp?Menu=" & sMenu & "&Nombre=" & sNombre
    Celda   = ofv.BotonVolver(sAccion)     
  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no","","",celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>