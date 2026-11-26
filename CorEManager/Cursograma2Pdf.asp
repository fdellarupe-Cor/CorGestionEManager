<%


ncodigo=request.querystring("codigo")
stipo=request.querystring("tipo")




%>
<HTML>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=Edge" />
</head>
<style>
  .rotateDiv	{
	   		   			margin-top:300px;
	   		   			-webkit-transform:rotate(120deg);
	   		   			-moz-transform:rotate(120deg);
	   		   			-o-transform:rotate(120deg);
	   		   			/* filter:progid:DXImageTransform.Microsoft.BasicImage(rotation=1.5); */
	   		   			-ms-transform:rotate(270deg);

		   		}
	 .drsElement {
	  position: absolute;
	  border: 1px solid #333;
	 }



       .block { position: absolute; }


.dd1{

                     border-left-style:solid;
                     border-left-width:1px;
                     border-left-color:#99ccff;
                     border-bottom-style:solid;
                     border-bottom-width:1px;
                     border-bottom-color:#99ccff;
                     font-size:8px;
		     font-face:arial;
}

.dd{
        border: 1px solid #99ccff; }


       .container_2 {

                     font-size:10px;
		     font-face:arial;
		     background: #f7eedd;



}

.dv{
        border: none; }


       #container_4 {

                     font-size:12px;
		     font-face:arial;


}


.dx{
        border: 1px solid #eeeeee; }


       #container_5 {

                     font-size:10px;
		     font-face:arial;
		     background: #F7BE81;

                     }
</style>
<body xBGCOLOR="#ff0000">
<%

if  SESSION("FILETIPO")  = "L" then
  aspf="Cursograma2Pdf2.asp"
else
  aspf="fileserver/asp/Cursograma2Pdf2.asp"
end if
%>

<form enctype="multipart/form-data" style="visibility:hidden;" id="genpdf"  method=post action="<%=aspf%>?codigo=<%=ncodigo%>&tipo=<%=stipo%>"   >
 <input type="text" id="spdf" name="FILE1">
 <input type=submit>





</form>
 <button id='bt1' style='visibility:hidden;' onclick="generateImages(1060,700)">generar Img</button>
 <button id='bt2' style='visibility:hidden;' onclick="genpdf()">generar pdf</button>
 <div id="myImages" ></div>

<div id=zzf  >
<%
xx=request.form("fname")
xx=  xx & "<script type='text/javascript' src='JSGRAF/html2canvas.js'></script>"
xx=  xx & "<script type='text/javascript' src='JSGRAF/jspdf.js'></script>"
xx=  xx & "<script type='text/javascript' src='JSGRAF/jspdf.plugin.addimage.js'></script> "
xx=  xx & " <script type='text/javascript' src='JSGRAF/jspdf.source.js'></script>"

response.write xx

%>
</div>
<script>
  var xid=0;
  var strind;

  function asignar(xthis)
  	{

            BreakItUp(xthis);

  	  return false;
	}
	 function getElementsByClassName(node, classname) {
			     var a = [];
			     var re = new RegExp('(^| )'+classname+'( |$)');
			     var els = node.getElementsByTagName("*");
			     for(var i=0,j=els.length; i<j; i++)
			         if(re.test(els[i].className))a.push(els[i]);
			     return a;
}
function drawCanvas(texto,imgid) {

	 		  		   var canvas = document.createElement('canvas');
	 		  		   canvas.width='15'
	 		  		   canvas.height='295';
	                     var img = document.getElementById(imgid);
	                   //  alert(img);
	 		  		   var context = canvas.getContext('2d');

	 		  		 // Fill Stroke Text
	                    var x = canvas.width / 2;
	       				var y = canvas.height / 2;
	 		  		   context.lineWidth=1;

	 		  		   context.strokeStyle="rgba(35, 70, 237, .8)";
	                   //  context.translate(70 / 2, 40 / 2);
	 		  		   context.font="10pt arial";
	 		  		    context.translate(x ,y);
	                    context.rotate(270*Math.PI/180);

	                    context.textAlign = 'center';
	 		  		   context.fillText(texto, 0, 4);

	 		  		//   context.strokeText("Fill, Stroke Text, 28pt, sans-serif", 0, 0);


	          img.src=canvas.toDataURL("")
	 		  		 }
function pep()
     {
              var alto =document.getElementById('mainTabla').clientHeight+5;
               var largo = document.getElementById('mainTabla').clientWidth+5;
                 //document.getElementById('bt1').style.display='none';
              //   document.getElementById('bt2').style.display='none';
                  var ele=getElementsByClassName( document.body,'secox');

				      for(i=0;i< ele.length;i++)
				      {
				       title= document.getElementById (ele[i].id).title;
				       drawCanvas(title,ele[i].id);
				       document.getElementById (ele[i].id).style.display='';

}
				        var ele=getElementsByClassName( document.body,'spanxx');

					   				      for(i=0;i< ele.length;i++)
					   				      {
					   				     // alert(document.getElementById (ele[i].id).outerHTML);
					   				    document.getElementById (ele[i].id).outerHTML ='';



     }

	 	         html2canvas(  document.body , {
	 	           onrendered: function(canvas) {

	 	             document.body.appendChild(canvas);
	 	             canvas.id='xx11';

	 	            document.getElementById('zzf').style.display='none';
                   // document.getElementById('bt1').style.visibility='';
                   // document.getElementById('bt2').style.visibility='';
                    generateImages(1060,700);
                    genpdf();
                    document.getElementById('spdf').value=strind;
                  document.getElementById('genpdf').submit();


                //    alert(strind);
                //    GotPDF(strind);



	 	           }, width: largo,
height: alto
	 	         });

	 	         }
	 	         pep();

function generateImages1(width, height) {

      width = parseInt(width);
      height = parseInt(height);
      if (isNaN(width)) width = 100;
      if (isNaN(height)) height = 100;
     width = Math.max(width, 50);
      height = Math.max(height, 50);

   //   alert(width);
    // document.getElementById('mainCanvas').style.display='none';
      var imgDiv = document.getElementById('myImages');
      imgDiv.innerHTML = ''; // clear out the old images, if any
     // var db = myDiagram.documentBounds.copy();
     var canvas = document.getElementById('xx11');
     var context = canvas.getContext('2d');

      var boundswidth = canvas.width;
      var boundsheight = canvas.height;
      var imgWidth = width;
      var imgHeight = height;
   //   var p = db.position.copy();
      for (var i = 0; i < boundsheight; i += imgHeight) {
        for (var j = 0; j < boundswidth; j += imgWidth) {

         var xcanvas = document.createElement('canvas');
         xcanvas.width=width;
         xcanvas.height=height;
       //  alert('1');
         var context2 = xcanvas.getContext('2d');
         context2.scale(0.1,0.1);

         var imgData=context.getImageData(j,i,width,height);
         //
           context2.lineWidth=1;

		  		 		  		   context2.strokeStyle="rgba(35, 70, 237, .8)";

		  		 		  		   context2.font="10pt arial";


		  		                    context2.textAlign = 'center';
		 		  		   context2.fillText('11111111111111111', 0, 4);
         context2.putImageData(imgData,0,0);
        // context2.drawImage( imgData, 0, 0, 2, 2 );

        //  alert(j + ' ' + i);
          // Append the new HTMLImageElement to the #myImages div
        //  img.className = 'images';
         imgDiv.appendChild(xcanvas);
         imgDiv.appendChild(document.createElement('br'));
        }
      }
document.getElementById('xx11').style.display='none';


    }


    function drawCanvasf(texto) {

		 		  		   var canvas = document.createElement('canvas');
		 		  		   canvas.width='15';
		 		  		   canvas.height='15';

		 		  		   var context = canvas.getContext('2d');



		 		  		   context.lineWidth=1;

		 		  		   context.strokeStyle="rgba(35, 70, 237, .8)";

		 		  		   context.font="10pt arial";


		                    context.textAlign = 'center';
		 		  		   context.fillText(texto, 0, 4);



              return canvas;

	 		  		 }

function generateImages(width, height) {

      width = parseInt(width);
      height = parseInt(height);
      if (isNaN(width)) width = 100;
      if (isNaN(height)) height = 100;
     width = Math.max(width, 50);
      height = Math.max(height, 50);
      document.getElementById('bt1').style.display='none';
      document.getElementById('bt2').style.display='none';

   //     alert(width + '  ' + height);
    // document.getElementById('mainCanvas').style.display='none';
      var imgDiv = document.getElementById('myImages');
      imgDiv.innerHTML = ''; // clear out the old images, if any
     // var db = myDiagram.documentBounds.copy();
     var canvas = document.getElementById('xx11');
     var context = canvas.getContext('2d');

      var boundswidth = canvas.width;
      var boundsheight = canvas.height;
      var imgWidth = width;
      var imgHeight = height;
      var pagx=0
      var pagy=0

   //   var p = db.position.copy();
      for (var i = 0; i < boundsheight; i += imgHeight) {
        pagx++;
        pagy=0;
        for (var j = 0; j < boundswidth; j += imgWidth) {
         xid++;
         pagy++
         var xcanvas = document.createElement('canvas');
         xcanvas.width=width;
         xcanvas.style.visibility='hidden';
         xcanvas.height=height+50;
         xcanvas.id='ppp' + xid;
       //  alert('1');
         var context2 = xcanvas.getContext('2d');
    // context2.fillStyle = "#ffffff";
       //   context2.fillRect(0,0,width,height);
         var imgData=context.getImageData(j,i,width,height);
         //   context2.clearRect(0, 0, width,height);
          context2.putImageData(imgData,0,0);
   //       context2.globalCompositeOperation = 'source-atop';
	//	      context2.fillStyle = 'rgba(0,0,0,.5)';
    //      context2.fillRect(0,0,width,height);
              context2.lineWidth=1;

			context2.strokeStyle="rgba(35, 70, 237, .8)";

			context2.font="10pt arial";


			context2.textAlign = 'left';
		   context2.fillText('Página ' + pagx + '/' + pagy  ,10, height+ 45);

        //  alert(j + ' ' + i);
          // Append the new HTMLImageElement to the #myImages div
        //  img.className = 'images';
         imgDiv.appendChild(xcanvas);
         imgDiv.appendChild(document.createElement('br'));


        // break
        }
       // break

      }
document.getElementById('xx11').style.display='none';

 document.getElementById('bt1').style.display='';
      document.getElementById('bt2').style.display='';


    }


    function genpdf()
    {



      var doc = new jsPDF('landscape')

       for (i=1;i<= xid;i++)
       {

       var xcanvas =document.getElementById('ppp'+ i);
       var imgData = canvasToImage(xcanvas, xcanvas.width,xcanvas.height); //xcanvas.toDataURL("image/jpeg");
     //    alert(i + ' ' + imgData);
	     doc.addImage(imgData, 'JPEG',0, 0,0,0);
       // if(i +1 < xid  ) { doc.addPage();}

if (i !=xid){         doc.addPage();}

       }

   //doc.output('datauristring');
//doc.output('save','pep2.pdf');
     //    doc.save('Test.pdf');
        strind = doc.output('datauristring');




//GotPDF(strind);
	//   	     var x = window.open();
 // x.document.open();
 // x.document.location=strind;

    }

function GotPDF(data)
{

   var  a = document.createElement("a");
    a.href = data;
    a.download = "documentName.pdf"
    a.click();
}

function convertDataURIToBinary(dataURI) {
    var BASE64_MARKER = ';base64,';
  var base64Index = dataURI.indexOf(BASE64_MARKER) + BASE64_MARKER.length;
  var base64 = dataURI.substring(base64Index);
  var raw = window.atob(base64);
  var rawLength = raw.length;
  var array = new Uint8Array(new ArrayBuffer(rawLength));

  for(i = 0; i < rawLength; i++) {
    array[i] = raw.charCodeAt(i);
  }
  return array;
}
function canvasToImage(xcanvas,w,h)
{
//cache height and width


var data;

var context=xcanvas.getContext('2d');
//get the current ImageData for the canvas.
data = context.getImageData(0, 0, w, h);

//store the current globalCompositeOperation
var compositeOperation = context.globalCompositeOperation;

//set to draw behind current content
context.globalCompositeOperation = "destination-over";

//set background color
context.fillStyle = "#ffffff";

//draw background / rect on entire canvas
context.fillRect(0,0,w,h);


//get the image data from the canvas
var imageData = xcanvas.toDataURL("image/jpeg");


//clear the canvas
context.clearRect (0,0,w,h);

//restore it with original / cached ImageData
context.putImageData(data, 0,0);

//reset the globalCompositeOperation to what it was
context.globalCompositeOperation = compositeOperation;


//return the Base64 encoded data url string
return imageData;
}

function BreakItUp(ffff)
{
  //Set the limit for field size.
  var FormLimit = 102399;

  //Get the value of the large input object.
  var TempVar = new String;
  TempVar = strind;//document.getElementById("spdf").value;

  //If the length of the object is greater than the limit, break it
  //into multiple objects.
  if (TempVar.length > FormLimit)
  {
    document.getElementById("spdf").value = TempVar.substr(0, FormLimit);
alert(TempVar.length);
    TempVar = TempVar.substr(FormLimit);
alert(TempVar.length);
    while (TempVar.length > 0)
    {

      var objTEXTAREA = document.createElement("file");
    alert(TempVar.length);
      objTEXTAREA.name = "spdf";
      objTEXTAREA.value = TempVar.substr(0, FormLimit);

      objTEXTAREA.style.display = '¿none';
      ffff.appendChild(objTEXTAREA);

      TempVar = TempVar.substr(FormLimit);
    }
  }
  return false;
}


</script>
         <script>





</script>

</body>

<HTML>


