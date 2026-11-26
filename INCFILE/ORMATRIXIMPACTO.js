function Open_All_Row()
{
	var estado = document.all('General').innerText;
	var newEstado = '';
	if (estado == '(+)')  
	{
		document.all('General').innerText = '(-)';
		newEstado = 'Abrir';
   	}
	else
   	{
   		document.all('General').innerText = '(+)';
		newEstado = 'Cerrar';
   	}
   	var i = 0;
	var aux = 'Empezo';
  // 		alert('hola');
  //		var fi = 'fi' + i + '_0'
  	while(aux == 'Empezo')
	{
		var nombre = 'particular_' + i;
		if(document.getElementById(nombre)!=null)
		{
			i += 1;
		}
		else
		{
			aux = 'Termino';
		}
	}	
  
   	for(j = 0; j < i; j++)
   	{
   		Open_Particular_Row(j, newEstado);
  //			alert(j + ' , ' + i);
   	}
}



//****************************************************************************************************************
//
//****************************************************************************************************************

function Open_Particular_Row(nrow1, estadoGeneral)
{
	var cambiarEstadoParticular = '';
	var particular = document.all('particular_' + nrow1).innerText;
	//alert(particular);
	if(estadoGeneral == '' || estadoGeneral == null)
	{
		if(particular == '(+)')
		{
			cambiarEstadoParticular = 'Abrir';
		}
		else
		{
			cambiarEstadoParticular = 'Cerrar';
		}
	}
	else
	{
		if(estadoGeneral == 'Abrir')
		{
			cambiarEstadoParticular = 'Abrir';
		}
		else
		{
			cambiarEstadoParticular = 'Cerrar';
		}
	}
			
	OpenParticularFi(nrow1, cambiarEstadoParticular)
	OpenParticularXind(nrow1, cambiarEstadoParticular)

}



//****************************************************************************************************************
//
//****************************************************************************************************************
	
function OpenParticularFi(nrow1, cambiarEstadoParticular)
{
	var i_fi = 0;
	var fi = 'fi' + nrow1 + '_' + i_fi;
	if(document.getElementById(fi)==null)
	{
		i_fi += 1;
		var fi = 'fi' + nrow1 + '_' + i_fi;
		if(document.getElementById(fi)==null)
   		{
   			i_fi += 1;
   			var fi = 'fi' + nrow1 + '_' + i_fi;
   			if(document.getElementById(fi)==null)
   			{
   			}
   		}
   	}
	
	while(fi != '')
	{
		if(document.getElementById(fi)!=null)
		{
			var nombre = 'particular_' + nrow1;
			//alert(nombre);
			//var estado = document.getElementById(nombre).innerHTML;
			//alert(estado);
			if(cambiarEstadoParticular == 'Abrir')
			{
				document.all(nombre).innerText = '(-)';
				hide_row(fi, 'Abrir');
				i_fi += 1;
				fi = 'fi' + nrow1 + '_' + i_fi;
			}
			else
			{
				document.all(nombre).innerText = '(+)';
				hide_row(fi, 'Cerrar');
				i_fi += 1;
				fi = 'fi' + nrow1 + '_' + i_fi;
			}
			
		}
		else
		{
			fi='';
		}
	}
}


//****************************************************************************************************************
//
//****************************************************************************************************************

function OpenParticularXind(nrow1, cambiarEstadoParticular)
{
	var i_xind = 0;
	var xind = 'xind' + nrow1 + '_' + i_xind;
	if(document.getElementById(xind)==null)
	{
		i_xind = 1;
		xind = 'xind' + nrow1 + '_' + i_xind;
		if(document.getElementById(xind)==null)
		{
			i_xind = 2;
			xind = 'xind' + nrow1 + '_' + i_xind;
			if(document.getElementById(xind)==null)
			{
				xind= '';
			}
		}
	}
	
	while(xind != '')
	{
		if(document.getElementById(xind)!=null)
		{
			var nombre = 'particular_' + nrow1;
			//alert(nombre);
			//var estado = document.getElementById(nombre).innerHTML;
			//alert(estado);
			if(cambiarEstadoParticular == 'Abrir')
			{
				document.getElementById(nombre).innerHTML = '(-)';
				hide_row(xind, 'Abrir');
				i_xind += 1;
				xind = 'xind' + nrow1 + '_' + i_xind;
			}
			else
			{
				document.getElementById(nombre).innerHTML = '(+)';
				hide_row(xind, 'Cerrar');
				i_xind += 1;
				xind = 'xind' + nrow1 + '_' + i_xind;
			}
		}
		else
		{
			xind='';
		}
	}
}



//****************************************************************************************************************
//
//****************************************************************************************************************

function hide_row(nrow, cambiarEstadoParticular)
{    
	//alert('entro al js');
	//alert(cambiarEstadoParticular);
	//alert(nrow);
		
	var v = document.getElementById(nrow);
		if(cambiarEstadoParticular == '' || cambiarEstadoParticular == null)
	{
		if (v.style.display =='none')
		{
			//alert('entro a none y si particular');
			v.style.display ='block';
			document.all('interna_' + nrow).innerText = '(-)';
			//alert('interna_' + nrow);
		}
		else
		{
			//alert('entro a block y si particular');
			v.style.display ='none';
			document.all('interna_' + nrow).innerText = '(+)';
			//alert('interna_' + nrow);
		}
	}
	else
	{
		//alert(cambiarEstadoParticular);
		if(cambiarEstadoParticular == 'Abrir')
		{
			v.style.display ='block';
			document.all('interna_' + nrow).innerText = '(-)';
			//alert('interna_' + nrow);
		}
		else
		{
			v.style.display ='none';
			document.all('interna_' + nrow).innerText = '(+)';
			//alert('interna_' + nrow);
		}
	}
}