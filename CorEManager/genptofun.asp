<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  'response.write session("codSol")
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

'  sHTML = sHTML & "<HTML>"
'  sHTML = sHTML & "<BODY onload='window.top.ldfobj();' >"


         sHTML = OFV.MenuHeader("","")

  spuesto = request.querystring("puesto")




  xTrColor = "#FFFFFF"

  sImageva    = "<font face='corflaticon' size=2 color='#00aa00'><b>&#0252;</b></font>" & "&nbsp;&nbsp;"

  wkstylof = "aaa style='font-family:arial;font-size:6pt;color:#888888;' "

  BCOLORG = "#cccccc"
  BCOLORD = "#6699CC"

'''''''''''''''''''''''''''''

  wkstylo1 = " style='background-color:#ddddff;color:#aaaaaa;font-weight:bold;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "
  wkstylo2 = " style='background-color:#FeDDAA;color:#aaaaaa;font-weight:bold;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "
  wkstylo3 = " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' "


  set cn = ofv.conectar(ofv.strconn2)






  'response.write "<center>"

    Fila = ""





    CELDAD = ""
    CeldaD = CeldaD & ofv.GenCelda("","utI  ","","95%","","","","DESCRIPCION DE FUNCIONES DEL PUESTO","si")
    FilaD     = ofv.genrow("","","","","","","",CeldaD,"si")
    sFORMhed = ofv.GenTabla("","TWIND","","100%","","0","","0","1","","",FilaD,"si")


   shtml = shtml & sFORMhed


     fila=""


     sperfil = spuesto




  Strsql = "Select Ecoperfil.*,Tipoperfil.descri From Ecoperfil, Tipoperfil where "
  Strsql = Strsql & "Ecoperfil.codperfil = Tipoperfil.tipoperfil and Ecoperfil.OBJID = 0" & spuesto
'' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)



   if  rs.eof then


    Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

    Celda = ofv.GenCelda("","uti colspan=2 style='border-style:solid;border-color:#888888;border-width:1;'  ","center","","","","","<big>No hay funciones descriptas para el puesto.</big>","si")
    Fila = fila & ofv.genrow("","","","","50","","",Celda,"si")

 else

   Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","30%","","","","AREA/SECTOR:","si")


  SSECTOR = RS(1)
  SSECTORDAT = RS(1)

   Strsql = "Select * From ECO where eco = 0" & RS(1)
   set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rsX.EOF then
      SSECTOR = RSX(1)
   END IF

    'SECTOR
    sInput = SSECTOR
    Celda = Celda & ofv.gencelda("","utI ","","70%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")


    Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","PUESTO:","si")

    'Perfil
    sInput = rs(4)
    Celda = Celda & ofv.gencelda("","utI ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")

    Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","DOTACION:","si")

    'Dotacion
    sInput = rs(3)
    Celda = Celda & ofv.gencelda("","utI ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")



   end if


 shtml =shtml & "<center>" & ofv.GenTabla("","AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ","","95%   ","","0","","1","1","","",fila,"si") & "</center>"


  Formulario = ""

   Celda = ofv.GenCelda("","UTI COLSPAN=3  " & wkstylo1,"","30%","","","","DESCRIPCION DEL PUESTO","si")
 '  Celda = CELDA & ofv.GenCelda("","UTI COLSPAN=2  ","","","","","","&nbsp;","si")

   Fila  = ofv.genrow("","","","","25","","",Celda,"si")



  Strsql = "Select * FROM ECOPERFILDSC WHERE "
  Strsql = Strsql & " ECOPERFIL = 0" & sperfil & " ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)





  posI = 0
   IF Not rs.EOF THEN

   Formulario = Formulario & FILA

  Do While (Not rs.EOF)

    Fila = ""


    sinput = "sin definir"
    Strsql = "Select * From TIPODSCFUN where OBJID = 0" & rs(3)
    set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsx2.eof then sinput = rsx2(1)



   Celda = ofv.GenCelda("","UTI   " & wkstylo2,"","30%","","","",sinput,"si")
   Celda = CELDA & ofv.GenCelda("","UTI " & wkstylo3,"","","","","","&nbsp;","si")
   Celda = CELDA & ofv.GenCelda("","UTI " & wkstylo3,"","","","","","&nbsp;","si")
   Fila  = Fila & ofv.genrow(""," ","","","25","","",Celda,"si")

  ''  sInput = sImageva & rs(4)
      sInput = sImageva & replace(rs(4) & "",chr(13),"<br>" & sImageva)

    Celda =  ofv.GenCelda("","uti COLSPAN=3  ","","","","","",sInput ,"si")
    Fila  = Fila & ofv.genrow("","","","","50","","",Celda,"si")

    Formulario = Formulario & Fila

    posi = posi + 1
    rs.Movenext
  Loop

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ","","95%","","0","","1","1","","",Formulario,"si") & "</center>"

 ELSE

 '   Formulario = Formulario & FILA
 '   sInput = " SIN DEFINIR DESCRIPCIONES PARA EL PUESTO. "
 '   Celda =  ofv.GenCelda("","uti COLSPAN=3  ","","","","","",sInput ,"si")
 '   Fila = ofv.genrow("","","","","50","","",Celda,"si")
 '   Formulario = Formulario &  FILA

  END IF







    Formulario = ""

  Strsql = "Select * FROM ECOPERFILDEP WHERE "
  Strsql = Strsql & " ECOPERFIL = 0" & sperfil & " AND PTOTIPO ='RA' ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Fila = ""




   Celda = ofv.GenCelda("","UTI COLSPAN=2  " & wkstylo1,"","","","","","REPORTA A","si")
   Fila  = ofv.genrow("","","","","25","","",Celda,"si")

   Celda = ""
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","50%","","","","AREA/SECTOR","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","50%","","","","PERFIL","si")
 '  Fila  = Fila & ofv.genrow("","","","","25","","",Celda,"si")



  posI = 0
   IF Not rs.EOF THEN

    Formulario = Formulario & FILA

  Do While (Not rs.EOF)

      Celda = ""


    'SECTOR
    sinput = "sin definir"
    sinput1 = ""
    sinput2 = ""

    Strsql = "Select * From ECO where eco = 0" & rs(4)
    set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsx2.eof then sinput1 = rsx2(1)


  ''''  Celda = Celda & ofv.gencelda("","uti  ","","60%","","","",sInput,"si")

    'Perfil
     sinput = "sin definir"
    Strsql = "Select * From TIPOPERFIL where TIPOPERFIL = '" & rs(5) & "' "
    set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsx2.eof then sinput2 = rsx2(1)

    if sinput1 <> "" or sinput2 <> "" then
       sinput = sImageva & sinput1 & "&nbsp;-&nbsp;" & sinput2
    end if

    Celda = Celda & ofv.gencelda("","uti colspan=2 ","","60%","","","",sInput,"si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")


    Formulario = Formulario & Fila


    posi = posi + 1
    rs.Movenext
  Loop

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ","","95%","","0","","1","1","","",Formulario,"si") & "</center>"

  ELSE

  '  Formulario = Formulario & FILA
  '  sInput = " SIN DEFINIR SECTOR/PUESTO A QUIEN REPORTA. "
  '  Celda =  ofv.GenCelda("","uti COLSPAN=2  ","","","","","",sInput ,"si")
  '  Fila = ofv.genrow("","","","","50","","",Celda,"si")
  '  Formulario = Formulario &  FILA


  END IF










    Formulario = ""

  Strsql = "Select * FROM ECOPERFILDEP WHERE "
  Strsql = Strsql & " ECOPERFIL = 0" & sperfil & " AND PTOTIPO ='RB' ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Fila = ""




   Celda = ofv.GenCelda("","UTI COLSPAN=2  " & wkstylo1,"","","","","","PUESTO QUE LE REPORTA","si")
   Fila  = ofv.genrow("","","","","25","","",Celda,"si")

   Celda = ""
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","50%","","","","AREA/SECTOR","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","50%","","","","PERFIL","si")
 '  Fila  = Fila & ofv.genrow("","","","","25","","",Celda,"si")



  posI = 0
   IF Not rs.EOF THEN

    Formulario = Formulario & FILA

  Do While (Not rs.EOF)


      Celda = ""


    'SECTOR
    sinput = "sin definir"
    sinput1 = ""
    sinput2 = ""

    Strsql = "Select * From ECO where eco = 0" & rs(4)
    set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsx2.eof then sinput1 = rsx2(1)


  ''''  Celda = Celda & ofv.gencelda("","uti  ","","60%","","","",sInput,"si")

    'Perfil
     sinput = "sin definir"
    Strsql = "Select * From TIPOPERFIL where TIPOPERFIL = '" & rs(5) & "' "
    set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsx2.eof then sinput2 = rsx2(1)

    if sinput1 <> "" or sinput2 <> "" then
       sinput = sImageva & sinput1 & "&nbsp;-&nbsp;" & sinput2
    end if

    Celda = Celda & ofv.gencelda("","uti colspan=2 ","","60%","","","",sInput,"si")


    Fila = ofv.genrow("","","","","25","","",Celda,"si")


    Formulario = Formulario & Fila



    posi = posi + 1
    rs.Movenext
  Loop

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ","","95%","","0","","1","1","","",Formulario,"si") & "</center>"

  ELSE

 '   Formulario = Formulario & FILA
 '   sInput = " SIN DEFINIR SECTOR/PUESTO QUE LE REPORTA. "
 '   Celda =  ofv.GenCelda("","uti COLSPAN=2  ","","","","","",sInput ,"si")
 '   Fila = ofv.genrow("","","","","50","","",Celda,"si")
 '   Formulario = Formulario &  FILA


  END IF





    Formulario = ""

  Strsql = "Select * FROM ECOPERFILFUN WHERE TIPO <> 'G' AND "
  Strsql = Strsql & " ECOPERFIL = 0" & sperfil & " ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Fila = ""

   Celda = ofv.GenCelda("","UTI COLSPAN=2  " & wkstylo1,"","","","","","FUNCIONES DEL PUESTO","si")

   Fila  = ofv.genrow("","","","","25","","",Celda,"si")

   Celda = ""
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","30%","","","","FUNCION","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","70%","","","","DESCRIPCION","si")
 '  Fila  = Fila & ofv.genrow("","","","","25","","",Celda,"si")




  posI = 0
   IF Not rs.EOF THEN

    Formulario = Formulario & FILA

  Do While (Not rs.EOF)

    sForm = "formRB" & posi


      Celda = ""
      Fila  = ""

    'ACCION
    sinput = "sin definir"
    Strsql = "Select * From TIPOACCIONFUN where objid = 0" & rs(3)
    set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsx2.eof then sinput = rsx2(1)
 '''   Celda = Celda & ofv.gencelda("","uti  ","","","","","",sInput,"si")

   Celda = ofv.GenCelda("","UTI   " & wkstylo2,"","30%","","","",sinput,"si")
   Celda = CELDA & ofv.GenCelda("","UTI " & wkstylo3,"","","","","","&nbsp;","si")
   Fila  = Fila & ofv.genrow(""," ","","","25","","",Celda,"si")


    'DESCRIPCION
   '''' sInput = sImageva & rs(4)
     sInput = sImageva & replace(rs(4) & "",chr(13),"<br>" & sImageva)
     Celda = ofv.gencelda("","uti colspan=2 ","","","","","",sInput,"si")
    Fila  = Fila & ofv.genrow("","","","","50","","",Celda,"si")




    Formulario = Formulario & Fila



    posi = posi + 1
    rs.Movenext
  Loop

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ","","95%","","0","","1","1","","",Formulario,"si") & "</center>"

  ELSE

  '  Formulario = Formulario & FILA
  '  sInput = " SIN DEFINIR FUNCIONES DEL PUESTO. "
  '  Celda =  ofv.GenCelda("","uti COLSPAN=2  ","","","","","",sInput ,"si")
  '  Fila = ofv.genrow("","","","","50","","",Celda,"si")
  '  Formulario = Formulario &  FILA


  END IF




    Formulario = ""

  Strsql = "Select * FROM ECOPERFILCOM WHERE "
  Strsql = Strsql & " ECOPERFIL = 0" & sperfil & " ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Fila = ""

   Celda = ofv.GenCelda("","UTI COLSPAN=3  " & wkstylo1,"","30%","","","","COMPETENCIAS DEL PUESTO","si")
 '  Celda = CELDA & ofv.GenCelda("","UTI COLSPAN=2  ","","","","","","&nbsp;","si")

   Fila  = ofv.genrow("","","","","25","","",Celda,"si")


  posI = 0
   IF Not rs.EOF THEN

   Formulario = Formulario & FILA

  Do While (Not rs.EOF)

      Celda = ""
      Fila  = ""

       SCOMPEDAT = "SIN DEFINIR"

    Strsql = "Select * From tipocompetencia where objid = 0" & rs(3)
'' response.write strsql
    set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    if  not rsX2.eof then
        SCOMPETIPO = CSTR(rsX2(2) & "")
        SCOMPEDAT  = CSTR(rsX2(1) & "")
    END IF

      sInput = SCOMPEDAT

   Celda = ofv.GenCelda("","UTI   " & wkstylo2,"","30%","","","",sinput,"si")
   Celda = CELDA & ofv.GenCelda("","UTI " & wkstylo3,"","","","","","&nbsp;","si")
   Celda = CELDA & ofv.GenCelda("","UTI " & wkstylo3,"","","","","","&nbsp;","si")
   Fila  = Fila & ofv.genrow(""," ","","","25","","",Celda,"si")


    'DESCRIPCION


    IF SCOMPETIPO = "1" THEN

       SCOMPEDAT = "<FONT COLOR='#993333' >SIN DEFINIR</FONT>"
       SCOMPEDET = rs(4) & ""

       IF NOT ISNUMERIC(SCOMPEDET) THEN SCOMPEDET = 0

      Strsql = "Select * From TIPOCOMPEDET where objid = 0" & SCOMPEDET
'' response.write strsql
      set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      if  not rsX2.eof then
        SCOMPEDAT = CSTR(rsX2(2) & "")
      END IF


    ELSE

        SCOMPEDAT = CSTR(rs(4) & "")

   END IF


      sInput = sImageva & replace(SCOMPEDAT & "",chr(13),"<br>" & sImageva)

    Celda =  ofv.GenCelda("","uti COLSPAN=3  ","","","","","",sInput ,"si")
    Fila  = Fila & ofv.genrow("","","","","50","","",Celda,"si")

    Formulario = Formulario & Fila



    posi = posi + 1
    rs.Movenext
  Loop

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","AAA style='border-style:solid;border-color:#cccccc;border-width:1;'  ","","95%","","0","","1","1","","",Formulario,"si") & "</center>"

  ELSE

  '  Formulario = Formulario & FILA
  '  sInput = " SIN DEFINIR COMPETENCIAS DEL PUESTO. "
  '  Celda =  ofv.GenCelda("","uti COLSPAN=3   ","","","","","",sInput ,"si")
  '  Fila = ofv.genrow("","","","","50","","",Celda,"si")
  '  Formulario = Formulario &  FILA


  END IF





  Call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</BODY>"

  sHTML = sHTML & "</HTML>"
end if

 response.write shtml



%>