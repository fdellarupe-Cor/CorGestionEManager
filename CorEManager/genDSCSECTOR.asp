<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<% 
   
ok=ofv.CheckUsuario() 
if ok <> "Y" then
  uv = ofv.uv 
  if uv = "NO" then ok = ofv.ValidarUsuario 
  sHTML = ok    
else     
  response.expires=0         
   
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar    
  if session("Nomodificar") = "S" then bModificar = False   

'  sHTML = sHTML & "<HTML>"
'  sHTML = sHTML & "<BODY onload='window.top.ldfobj();' >"
  

         sHTML = OFV.MenuHeader("","")

  spuesto = request.querystring("SECTOR")    

  
 
    
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
    CeldaD = CeldaD & ofv.GenCelda("","utI  ","","95%","","","","DESCRIPCION DE FUNCIONES DEL SECTOR","si")
    FilaD     = ofv.genrow("","","","","","","",CeldaD,"si")
    sFORMhed = ofv.GenTabla("","TWIND","","100%","","0","","0","1","","",FilaD,"si")


   shtml = shtml & sFORMhed


     fila=""


     sperfil = spuesto      
 



  Strsql = "Select * FROM ECO where ECO = 0" & SPUESTO

'' response.write strsql  
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   
    
  
   if  rs.eof then


    Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

    Celda = ofv.GenCelda("","uti colspan=2 style='border-style:solid;border-color:#888888;border-width:1;'  ","center","","","","","<big>No hay funciones descriptas para el Sector.</big>","si")
    Fila = fila & ofv.genrow("","","","","50","","",Celda,"si")

 else

   Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","30%","","","","AREA/SECTOR:","si")


  SSECTOR = RS(1)
  SSECTORDAT = RS(1)

  SSECTORSG = "----"
  SSECTORJR = "----"
  SSECTORNJ = "----" 


  sSQL = "Select * From SEGMENTO WHERE CODSEG = 0" & RS(4)
  Set rs3 = OFV.crearconsulta2(sSQL, cn)
  If NOT rs3.EOF Then  SSECTORSG = RS3(1)  


  sSQL = "Select * From JERARQUIA WHERE CODjerar = '" & RS(6) & "' "
  Set rs3 = OFV.crearconsulta2(sSQL, cn)
  If NOT rs3.EOF Then  SSECTORJR = RS3(1)  

  sSQL = "Select * From NIVELJERAR WHERE CODnivel = 0" & RS(5)
  Set rs3 = OFV.crearconsulta2(sSQL, cn)
  If NOT rs3.EOF Then   SSECTORNJ = RS3(1) 

    'SECTOR
    sInput = SSECTOR
    Celda = Celda & ofv.gencelda("","utI ","","70%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")


    Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","CODIGO:","si")

    'CODIGO
    sInput = rs(2)
    Celda = Celda & ofv.gencelda("","utI ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")


    Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","SEGMENTO:","si")

    'SEGMENTO
    sInput = SSECTORSG
    Celda = Celda & ofv.gencelda("","utI ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")

    Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","JERARQUIA:","si")

    'JERARQUIA
    sInput = SSECTORjr
    Celda = Celda & ofv.gencelda("","utI ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")

    Celda = ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","","NIVEL:","si")

    'NIVEL
    sInput = SSECTORnj
    Celda = Celda & ofv.gencelda("","utI ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")


   end if
  

 shtml =shtml & "<center>" & ofv.GenTabla("","AAA style='border-style:solid;border-color:#cccccc;border-width:1;' ","","95%   ","","0","","1","1","","",fila,"si") & "</center>"

 
  Formulario = ""

   Celda = ofv.GenCelda("","UTI COLSPAN=3  " & wkstylo1,"","30%","","","","DESCRIPCION DEL SECTOR","si")
 '  Celda = CELDA & ofv.GenCelda("","UTI COLSPAN=2  ","","","","","","&nbsp;","si")

   Fila  = ofv.genrow("","","","","25","","",Celda,"si")



  Strsql = "Select * FROM ECODESCRIPCION WHERE ECO = 0" & SPUESTO
  Strsql = Strsql & "  ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)





  posI = 0
   IF Not rs.EOF THEN

   Formulario = Formulario & FILA

  Do While (Not rs.EOF)

    Fila = ""


    sinput = "sin definir"
    Strsql = "Select * From TIPODSCSECTOR where OBJID = 0" & rs(3)
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


     






  Call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</BODY>"

  sHTML = sHTML & "</HTML>"
end if

 response.write shtml



%>    