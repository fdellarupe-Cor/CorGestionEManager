<%

  if request.querystring("sigo") <> "" then
     ssigo = Ucase(request.querystring("sigo"))
  else
     ssigo = "000000"
  end if

  ssigoa = ssigo


   redim preserve shex(16)
   shex(0) = "0"
   shex(1) = "1"
   shex(2) = "2"
   shex(3) = "3"
   shex(4) = "4"
   shex(5) = "5"
   shex(6) = "6"
   shex(7) = "7"
   shex(8) = "8"
   shex(9) = "9"
   shex(10) = "A"
   shex(11) = "B"
   shex(12) = "C"
   shex(13) = "D"
   shex(14) = "E"
   shex(15) = "F"
   shex(16) = "0"

   redim preserve sh(6)

   if ssigoa <> "000000" then
   sh(1) = left(ssigoa,1)
   sh(2) = mid(ssigoa,2)
   sh(2) = left(sh(2),1)
   sh(3) = mid(ssigoa,3)
   sh(3) = left(sh(3),1)
   sh(4) = mid(ssigoa,4)
   sh(4) = left(sh(4),1)
   sh(5) = mid(ssigoa,5)
   sh(5) = left(sh(5),1)
   sh(6) = right(ssigoa,1)
   ist = 4
   do until ist < 1
   if sh(ist) = "0" then
      sh(ist) = "F"
   else
      istt = 1
      do until istt > 16
         if sh(ist) = shex(istt) then
            sh(ist) = shex(istt - 1)
            exit do
         end if                  
       istt = istt + 1  
      loop
      exit do
   end if 
   ist = ist - 1
   loop
  ssigop = sh(1) & sh(2) & sh(3) & sh(4) & "00"
   end if

'rojo mas
   sh(1) = left(ssigoa,1)
   sh(2) = mid(ssigoa,2)
   sh(2) = left(sh(2),1)
   sh(3) = mid(ssigoa,3)
   sh(3) = left(sh(3),1)
   sh(4) = mid(ssigoa,4)
   sh(4) = left(sh(4),1)
   sh(5) = mid(ssigoa,5)
   sh(5) = left(sh(5),1)
   sh(6) = right(ssigoa,1)
  if sh(1) = "F" and sh(2) = "F" then
   ssigorm = ssigoa
  else 
   ist = 2
   do until ist < 1
   if sh(ist) = "F" then
      sh(ist) = "0"
   else
      istt = 0
      do until istt > 16
         if sh(ist) = shex(istt) then
            sh(ist) = shex(istt + 1)
            exit do
         end if                  
       istt = istt + 1  
      loop
      exit do
   end if 
   ist = ist - 1
   loop
  ssigorm = sh(1) & sh(2) & sh(3) & sh(4) & "00"
   end if

'rojo menos
   sh(1) = left(ssigoa,1)
   sh(2) = mid(ssigoa,2)
   sh(2) = left(sh(2),1)
   sh(3) = mid(ssigoa,3)
   sh(3) = left(sh(3),1)
   sh(4) = mid(ssigoa,4)
   sh(4) = left(sh(4),1)
   sh(5) = mid(ssigoa,5)
   sh(5) = left(sh(5),1)
   sh(6) = right(ssigoa,1)
  if sh(1) = "0" and sh(2) = "0" then
   ssigorn = ssigoa
  else 
   ist = 2
   do until ist < 1
   if sh(ist) = "0" then
      sh(ist) = "F"
   else
      istt = 1
      do until istt > 16
         if sh(ist) = shex(istt) then
            sh(ist) = shex(istt - 1)
            exit do
         end if                  
       istt = istt + 1  
      loop
      exit do
   end if 
   ist = ist - 1
   loop
  ssigorn = sh(1) & sh(2) & sh(3) & sh(4) & "00"
   end if



  sHTML = "<html><body>"

  sHTML = sHTML & "<table width=80% marginwidth=0 cellspacing=1"
  sHTML = sHTML & " cellpadding=1 >"

  spant = 1
  do until spant > 16
  slin = 1
  sHTML = sHTML & "<tr>"
  do until slin > 16
  if ssigo > "666666" then
     sfon = "#000000"
  else
     sfon = "#ffffff"       
  end if
    sHTML = sHTML & "<td bgcolor=#" & ssigo & " width=30 "
    sHTML = sHTML & "style=border-style:outset;border-width:thin; >"
    sHTML = sHTML & "<a href=selcolorsh.asp?sigo=" & ssigo & " >" 
    sHTML = sHTML & "<font face=tahoma size=1 color="
    sHTML = sHTML & sfon & " >" & ssigo & "</font></a></td>"

   sh(1) = left(ssigo,1)
   sh(2) = mid(ssigo,2)
   sh(2) = left(sh(2),1)
   sh(3) = mid(ssigo,3)
   sh(3) = left(sh(3),1)
   sh(4) = mid(ssigo,4)
   sh(4) = left(sh(4),1)
   sh(5) = mid(ssigo,5)
   sh(5) = left(sh(5),1)
   sh(6) = right(ssigo,1)

'   response.write " ssigo: " & ssigo
'   response.write " conv: " & sh(1) & sh(2) & sh(3) & sh(4) & sh(5) & sh(6)
   ist = 6
   do until ist < 1
   if sh(ist) = "F" then
      sh(ist) = "0"
   else
      istt = 0
      do until istt > 16
         if sh(ist) = shex(istt) then
            sh(ist) = shex(istt + 1)
            exit do
         end if                  
       istt = istt + 1  
      loop
      exit do
   end if 
   ist = ist - 1
   loop
  ssigo = sh(1) & sh(2) & sh(3) & sh(4) & sh(5) & sh(6)
  slin = slin + 1 
  loop
  spant = spant + 1
  sHTML = sHTML & "</tr>"
  loop
  sHTML = sHTML & "</table>"

  sHTML = sHTML & "<br><br><table marginwidth=0 cellspacing=1"
  sHTML = sHTML & " cellpadding=1 >"
  sHTML = sHTML & "<tr>"

  if ssigoa <> "000000" then
  if ssigop > "666666" then
     sfon = "#000000"
  else
     sfon = "#ffffff"       
  end if
    sHTML = sHTML & "<td bgcolor=#" & ssigop & " width=30 "
    sHTML = sHTML & "style=border-style:outset;border-width:thin; >"
    sHTML = sHTML & "<a href=selcolor.asp?sigo=" & ssigop & ">" 
    sHTML = sHTML & "<font face=tahoma size=1 color="
    sHTML = sHTML & sfon & " >" & ssigop & "</font></a></td>"
  end if

  if ssigoa <> "ffff00" then
  if ssigo > "666666" then
     sfon = "#000000"
  else
     sfon = "#ffffff"       
  end if
    sHTML = sHTML & "<td bgcolor=#" & ssigo & " width=30 "
    sHTML = sHTML & "style=border-style:outset;border-width:thin; >"
    sHTML = sHTML & "<a href=selcolor.asp?sigo=" & ssigo & ">" 
    sHTML = sHTML & "<font face=tahoma size=1 color="
    sHTML = sHTML & sfon & " >" & ssigo & "</font></a></td>"
  end if

    sHTML = sHTML & "<td bgcolor=#ff0000 width=30 "
    sHTML = sHTML & "style=border-style:outset;border-width:thin; >"
    sHTML = sHTML & "<a href=selcolor.asp?sigo=ff0000 >" 
    sHTML = sHTML & "<font face=tahoma size=1 color=#ffffff"
    sHTML = sHTML & " >rojo</font></a></td>"

  if ssigorm > "ff6666" then
     sfon = "#000000"
  else
     sfon = "#ffffff"       
  end if
    sHTML = sHTML & "<td bgcolor=#" & ssigorm & " width=30 "
    sHTML = sHTML & "style=border-style:outset;border-width:thin; >"
    sHTML = sHTML & "<a href=selcolor.asp?sigo=" & ssigorm & ">" 
    sHTML = sHTML & "<font face=tahoma size=1 color="
    sHTML = sHTML & sfon & " >" & ssigorm & "</font></a></td>"

  if ssigorn > "ff6666" then
     sfon = "#000000"
  else
     sfon = "#ffffff"       
  end if
    sHTML = sHTML & "<td bgcolor=#" & ssigorn & " width=30 "
    sHTML = sHTML & "style=border-style:outset;border-width:thin; >"
    sHTML = sHTML & "<a href=selcolor.asp?sigo=" & ssigorn & ">" 
    sHTML = sHTML & "<font face=tahoma size=1 color="
    sHTML = sHTML & sfon & " >" & ssigorn & "</font></a></td>"

    sHTML = sHTML & "<td bgcolor=#00ff00 width=30 "
    sHTML = sHTML & "style=border-style:outset;border-width:thin; >"
    sHTML = sHTML & "<a href=selcolor.asp?sigo=00ff00 >" 
    sHTML = sHTML & "<font face=tahoma size=1 color=#ffffff"
    sHTML = sHTML & " >verde</font></a></td>"

    sHTML = sHTML & "<td bgcolor=#0000ff width=30 "
    sHTML = sHTML & "style=border-style:outset;border-width:thin; >"
    sHTML = sHTML & "<a href=selcolor.asp?sigo=000000 >" 
    sHTML = sHTML & "<font face=tahoma size=1 color=#ffffff"
    sHTML = sHTML & " >azul</font></a></td>"


  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"
  sHTML = sHTML & "</body></html>"

Response.Write sHTML

%>