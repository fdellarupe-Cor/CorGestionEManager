
var searchhi_string = '';

var defaultDiacriticsRemovalMap = [
    {'base':'A', 'letters':/[\u0041\u24B6\uFF21\u00C0\u00C1\u00C2\u1EA6\u1EA4\u1EAA\u1EA8\u00C3\u0100\u0102\u1EB0\u1EAE\u1EB4\u1EB2\u0226\u01E0\u00C4\u01DE\u1EA2\u00C5\u01FA\u01CD\u0200\u0202\u1EA0\u1EAC\u1EB6\u1E00\u0104\u023A\u2C6F]/g},
    {'base':'AA','letters':/[\uA732]/g},
    {'base':'AE','letters':/[\u00C6\u01FC\u01E2]/g},
    {'base':'AO','letters':/[\uA734]/g},
    {'base':'AU','letters':/[\uA736]/g},
    {'base':'AV','letters':/[\uA738\uA73A]/g},
    {'base':'AY','letters':/[\uA73C]/g},
    {'base':'B', 'letters':/[\u0042\u24B7\uFF22\u1E02\u1E04\u1E06\u0243\u0182\u0181]/g},
    {'base':'C', 'letters':/[\u0043\u24B8\uFF23\u0106\u0108\u010A\u010C\u00C7\u1E08\u0187\u023B\uA73E]/g},
    {'base':'D', 'letters':/[\u0044\u24B9\uFF24\u1E0A\u010E\u1E0C\u1E10\u1E12\u1E0E\u0110\u018B\u018A\u0189\uA779]/g},
    {'base':'DZ','letters':/[\u01F1\u01C4]/g},
    {'base':'Dz','letters':/[\u01F2\u01C5]/g},
    {'base':'E', 'letters':/[\u0045\u24BA\uFF25\u00C8\u00C9\u00CA\u1EC0\u1EBE\u1EC4\u1EC2\u1EBC\u0112\u1E14\u1E16\u0114\u0116\u00CB\u1EBA\u011A\u0204\u0206\u1EB8\u1EC6\u0228\u1E1C\u0118\u1E18\u1E1A\u0190\u018E]/g},
    {'base':'F', 'letters':/[\u0046\u24BB\uFF26\u1E1E\u0191\uA77B]/g},
    {'base':'G', 'letters':/[\u0047\u24BC\uFF27\u01F4\u011C\u1E20\u011E\u0120\u01E6\u0122\u01E4\u0193\uA7A0\uA77D\uA77E]/g},
    {'base':'H', 'letters':/[\u0048\u24BD\uFF28\u0124\u1E22\u1E26\u021E\u1E24\u1E28\u1E2A\u0126\u2C67\u2C75\uA78D]/g},
    {'base':'I', 'letters':/[\u0049\u24BE\uFF29\u00CC\u00CD\u00CE\u0128\u012A\u012C\u0130\u00CF\u1E2E\u1EC8\u01CF\u0208\u020A\u1ECA\u012E\u1E2C\u0197]/g},
    {'base':'J', 'letters':/[\u004A\u24BF\uFF2A\u0134\u0248]/g},
    {'base':'K', 'letters':/[\u004B\u24C0\uFF2B\u1E30\u01E8\u1E32\u0136\u1E34\u0198\u2C69\uA740\uA742\uA744\uA7A2]/g},
    {'base':'L', 'letters':/[\u004C\u24C1\uFF2C\u013F\u0139\u013D\u1E36\u1E38\u013B\u1E3C\u1E3A\u0141\u023D\u2C62\u2C60\uA748\uA746\uA780]/g},
    {'base':'LJ','letters':/[\u01C7]/g},
    {'base':'Lj','letters':/[\u01C8]/g},
    {'base':'M', 'letters':/[\u004D\u24C2\uFF2D\u1E3E\u1E40\u1E42\u2C6E\u019C]/g},
    {'base':'N', 'letters':/[\u004E\u24C3\uFF2E\u01F8\u0143\u00D1\u1E44\u0147\u1E46\u0145\u1E4A\u1E48\u0220\u019D\uA790\uA7A4]/g},
    {'base':'NJ','letters':/[\u01CA]/g},
    {'base':'Nj','letters':/[\u01CB]/g},
    {'base':'O', 'letters':/[\u004F\u24C4\uFF2F\u00D2\u00D3\u00D4\u1ED2\u1ED0\u1ED6\u1ED4\u00D5\u1E4C\u022C\u1E4E\u014C\u1E50\u1E52\u014E\u022E\u0230\u00D6\u022A\u1ECE\u0150\u01D1\u020C\u020E\u01A0\u1EDC\u1EDA\u1EE0\u1EDE\u1EE2\u1ECC\u1ED8\u01EA\u01EC\u00D8\u01FE\u0186\u019F\uA74A\uA74C]/g},
    {'base':'OI','letters':/[\u01A2]/g},
    {'base':'OO','letters':/[\uA74E]/g},
    {'base':'OU','letters':/[\u0222]/g},
    {'base':'P', 'letters':/[\u0050\u24C5\uFF30\u1E54\u1E56\u01A4\u2C63\uA750\uA752\uA754]/g},
    {'base':'Q', 'letters':/[\u0051\u24C6\uFF31\uA756\uA758\u024A]/g},
    {'base':'R', 'letters':/[\u0052\u24C7\uFF32\u0154\u1E58\u0158\u0210\u0212\u1E5A\u1E5C\u0156\u1E5E\u024C\u2C64\uA75A\uA7A6\uA782]/g},
    {'base':'S', 'letters':/[\u0053\u24C8\uFF33\u1E9E\u015A\u1E64\u015C\u1E60\u0160\u1E66\u1E62\u1E68\u0218\u015E\u2C7E\uA7A8\uA784]/g},
    {'base':'T', 'letters':/[\u0054\u24C9\uFF34\u1E6A\u0164\u1E6C\u021A\u0162\u1E70\u1E6E\u0166\u01AC\u01AE\u023E\uA786]/g},
    {'base':'TZ','letters':/[\uA728]/g},
    {'base':'U', 'letters':/[\u0055\u24CA\uFF35\u00D9\u00DA\u00DB\u0168\u1E78\u016A\u1E7A\u016C\u00DC\u01DB\u01D7\u01D5\u01D9\u1EE6\u016E\u0170\u01D3\u0214\u0216\u01AF\u1EEA\u1EE8\u1EEE\u1EEC\u1EF0\u1EE4\u1E72\u0172\u1E76\u1E74\u0244]/g},
    {'base':'V', 'letters':/[\u0056\u24CB\uFF36\u1E7C\u1E7E\u01B2\uA75E\u0245]/g},
    {'base':'VY','letters':/[\uA760]/g},
    {'base':'W', 'letters':/[\u0057\u24CC\uFF37\u1E80\u1E82\u0174\u1E86\u1E84\u1E88\u2C72]/g},
    {'base':'X', 'letters':/[\u0058\u24CD\uFF38\u1E8A\u1E8C]/g},
    {'base':'Y', 'letters':/[\u0059\u24CE\uFF39\u1EF2\u00DD\u0176\u1EF8\u0232\u1E8E\u0178\u1EF6\u1EF4\u01B3\u024E\u1EFE]/g},
    {'base':'Z', 'letters':/[\u005A\u24CF\uFF3A\u0179\u1E90\u017B\u017D\u1E92\u1E94\u01B5\u0224\u2C7F\u2C6B\uA762]/g},
    {'base':'a', 'letters':/[\u0061\u24D0\uFF41\u1E9A\u00E0\u00E1\u00E2\u1EA7\u1EA5\u1EAB\u1EA9\u00E3\u0101\u0103\u1EB1\u1EAF\u1EB5\u1EB3\u0227\u01E1\u00E4\u01DF\u1EA3\u00E5\u01FB\u01CE\u0201\u0203\u1EA1\u1EAD\u1EB7\u1E01\u0105\u2C65\u0250]/g},
    {'base':'aa','letters':/[\uA733]/g},
    {'base':'ae','letters':/[\u00E6\u01FD\u01E3]/g},
    {'base':'ao','letters':/[\uA735]/g},
    {'base':'au','letters':/[\uA737]/g},
    {'base':'av','letters':/[\uA739\uA73B]/g},
    {'base':'ay','letters':/[\uA73D]/g},
    {'base':'b', 'letters':/[\u0062\u24D1\uFF42\u1E03\u1E05\u1E07\u0180\u0183\u0253]/g},
    {'base':'c', 'letters':/[\u0063\u24D2\uFF43\u0107\u0109\u010B\u010D\u00E7\u1E09\u0188\u023C\uA73F\u2184]/g},
    {'base':'d', 'letters':/[\u0064\u24D3\uFF44\u1E0B\u010F\u1E0D\u1E11\u1E13\u1E0F\u0111\u018C\u0256\u0257\uA77A]/g},
    {'base':'dz','letters':/[\u01F3\u01C6]/g},
    {'base':'e', 'letters':/[\u0065\u24D4\uFF45\u00E8\u00E9\u00EA\u1EC1\u1EBF\u1EC5\u1EC3\u1EBD\u0113\u1E15\u1E17\u0115\u0117\u00EB\u1EBB\u011B\u0205\u0207\u1EB9\u1EC7\u0229\u1E1D\u0119\u1E19\u1E1B\u0247\u025B\u01DD]/g},
    {'base':'f', 'letters':/[\u0066\u24D5\uFF46\u1E1F\u0192\uA77C]/g},
    {'base':'g', 'letters':/[\u0067\u24D6\uFF47\u01F5\u011D\u1E21\u011F\u0121\u01E7\u0123\u01E5\u0260\uA7A1\u1D79\uA77F]/g},
    {'base':'h', 'letters':/[\u0068\u24D7\uFF48\u0125\u1E23\u1E27\u021F\u1E25\u1E29\u1E2B\u1E96\u0127\u2C68\u2C76\u0265]/g},
    {'base':'hv','letters':/[\u0195]/g},
    {'base':'i', 'letters':/[\u0069\u24D8\uFF49\u00EC\u00ED\u00EE\u0129\u012B\u012D\u00EF\u1E2F\u1EC9\u01D0\u0209\u020B\u1ECB\u012F\u1E2D\u0268\u0131]/g},
    {'base':'j', 'letters':/[\u006A\u24D9\uFF4A\u0135\u01F0\u0249]/g},
    {'base':'k', 'letters':/[\u006B\u24DA\uFF4B\u1E31\u01E9\u1E33\u0137\u1E35\u0199\u2C6A\uA741\uA743\uA745\uA7A3]/g},
    {'base':'l', 'letters':/[\u006C\u24DB\uFF4C\u0140\u013A\u013E\u1E37\u1E39\u013C\u1E3D\u1E3B\u017F\u0142\u019A\u026B\u2C61\uA749\uA781\uA747]/g},
    {'base':'lj','letters':/[\u01C9]/g},
    {'base':'m', 'letters':/[\u006D\u24DC\uFF4D\u1E3F\u1E41\u1E43\u0271\u026F]/g},
    {'base':'n', 'letters':/[\u006E\u24DD\uFF4E\u01F9\u0144\u00F1\u1E45\u0148\u1E47\u0146\u1E4B\u1E49\u019E\u0272\u0149\uA791\uA7A5]/g},
    {'base':'nj','letters':/[\u01CC]/g},
    {'base':'o', 'letters':/[\u006F\u24DE\uFF4F\u00F2\u00F3\u00F4\u1ED3\u1ED1\u1ED7\u1ED5\u00F5\u1E4D\u022D\u1E4F\u014D\u1E51\u1E53\u014F\u022F\u0231\u00F6\u022B\u1ECF\u0151\u01D2\u020D\u020F\u01A1\u1EDD\u1EDB\u1EE1\u1EDF\u1EE3\u1ECD\u1ED9\u01EB\u01ED\u00F8\u01FF\u0254\uA74B\uA74D\u0275]/g},
    {'base':'oi','letters':/[\u01A3]/g},
    {'base':'ou','letters':/[\u0223]/g},
    {'base':'oo','letters':/[\uA74F]/g},
    {'base':'p','letters':/[\u0070\u24DF\uFF50\u1E55\u1E57\u01A5\u1D7D\uA751\uA753\uA755]/g},
    {'base':'q','letters':/[\u0071\u24E0\uFF51\u024B\uA757\uA759]/g},
    {'base':'r','letters':/[\u0072\u24E1\uFF52\u0155\u1E59\u0159\u0211\u0213\u1E5B\u1E5D\u0157\u1E5F\u024D\u027D\uA75B\uA7A7\uA783]/g},
    {'base':'s','letters':/[\u0073\u24E2\uFF53\u00DF\u015B\u1E65\u015D\u1E61\u0161\u1E67\u1E63\u1E69\u0219\u015F\u023F\uA7A9\uA785\u1E9B]/g},
    {'base':'t','letters':/[\u0074\u24E3\uFF54\u1E6B\u1E97\u0165\u1E6D\u021B\u0163\u1E71\u1E6F\u0167\u01AD\u0288\u2C66\uA787]/g},
    {'base':'tz','letters':/[\uA729]/g},
    {'base':'u','letters':/[\u0075\u24E4\uFF55\u00F9\u00FA\u00FB\u0169\u1E79\u016B\u1E7B\u016D\u00FC\u01DC\u01D8\u01D6\u01DA\u1EE7\u016F\u0171\u01D4\u0215\u0217\u01B0\u1EEB\u1EE9\u1EEF\u1EED\u1EF1\u1EE5\u1E73\u0173\u1E77\u1E75\u0289]/g},
    {'base':'v','letters':/[\u0076\u24E5\uFF56\u1E7D\u1E7F\u028B\uA75F\u028C]/g},
    {'base':'vy','letters':/[\uA761]/g},
    {'base':'w','letters':/[\u0077\u24E6\uFF57\u1E81\u1E83\u0175\u1E87\u1E85\u1E98\u1E89\u2C73]/g},
    {'base':'x','letters':/[\u0078\u24E7\uFF58\u1E8B\u1E8D]/g},
    {'base':'y','letters':/[\u0079\u24E8\uFF59\u1EF3\u00FD\u0177\u1EF9\u0233\u1E8F\u00FF\u1EF7\u1E99\u1EF5\u01B4\u024F\u1EFF]/g},
    {'base':'z','letters':/[\u007A\u24E9\uFF5A\u017A\u1E91\u017C\u017E\u1E93\u1E95\u01B6\u0225\u0240\u2C6C\uA763]/g}
];
var changes;
function removeDiacritics (str) {
    if(!changes) {
		//alert('1');
        changes = defaultDiacriticsRemovalMap;
    }
    for(var i=0; i<changes.length; i++) {
        str = str.replace(changes[i].letters, changes[i].base);
    }
    return str;
}
function stripVowelAccent(str)
{
	var rExps=[ /[\xC0-\xC2]/g, /[\xE0-\xE2]/g,
		/[\xC8-\xCA]/g, /[\xE8-\xEB]/g,
		/[\xCC-\xCE]/g, /[\xEC-\xEE]/g,
		/[\xD2-\xD4]/g, /[\xF2-\xF4]/g,
		/[\xD9-\xDB]/g, /[\xF9-\xFB]/g ];

	var repChar=['A','a','E','e','I','i','O','o','U','u'];

	for(var i=0; i<rExps.length; ++i)
		str=str.replace(rExps[i],repChar[i]);

	return str;
}

function getText(elem) {
    if ((elem.textContent) && (typeof (elem.textContent) != "undefined")) {
        return elem.textContent.toLowerCase() ;
    } else {
        return elem.innerText.toLowerCase();
    }
}

function replaceNbsps(str) {
  var re = new RegExp(String.fromCharCode(160), "g");
  return str.replace(re, " ");
}

function highlightWord(node,word,doc) {
     doc = typeof(doc) != 'undefined' ? doc : document;
     count_matches = 0;
	// Iterate into this nodes childNodes

	//document.write(getText(node) );



	if (node.hasChildNodes) {
		var hi_cn;
		for (hi_cn=0;hi_cn<node.childNodes.length;hi_cn++) {
			count_matches = count_matches + highlightWord(node.childNodes[hi_cn],word,doc);
		}
	}

	// And do this node itself
	if (node.nodeType == 3) { // text node
		//tempNodeVal = stripVowelAccent(node.nodeValue.toLowerCase());
		//tempWordVal = stripVowelAccent(word.toLowerCase());
		//indexOf(str) === 0;
		tempNodeVal=removeDiacritics(node.nodeValue.toLowerCase());

		//tempWordVal =word;//removeDiacritics(word.toLowerCase());
		tempWordVal =removeDiacritics(word.toLowerCase());
    //    alert(tempNodeVal);
   //    alert(tempWordVal);
     tempWordVal=tempWordVal.replace(/\s+/gi,' ')

     tempWordVal=tempWordVal.replace(' ','\\s{0,5}');

    var number ="\\b" + tempWordVal + "\\b";
  // alert(number);
	var tt=new RegExp(number,"gi");
	var match =tt.exec(tempNodeVal);
              if (match==null)
              {
				  ni=-1;
     		  }
			else
			{
			 ni=match.index;
		    }



	if (ni !== -1) {
			pn = node.parentNode;
			if (pn.className != "searchword") {
                    count_matches = count_matches + 1;
				// word has not already been highlighted!
			//	alert( node.nodeValue);
			//	alert(tempNodeVal);
				nv = node.nodeValue;
			//	ni = tempNodeVal.indexOf(tempWordVal);
				//var number ="\\b" + tempWordVal + "\\b";
				//var tt=new RegExp(number,"gi");
              //  var match =tt.exec(tempNodeVal);
              //  ni=match.index;
               // alert(nv);
				// Create a load of replacement nodes
			//	ni = tempNodeVal.indexOf(tempWordVal);
			    var wlen=match[0].length;
				before = doc.createTextNode(nv.substr(0,ni));
				docWordVal = nv.substr(ni,wlen);
				after = doc.createTextNode(nv.substr(ni+wlen));
				hiwordtext = doc.createTextNode(docWordVal);
				hiword = doc.createElement("span");
				hiword.id='sxxxs';
				//hiword.style.zIndex="999";

				//hiword.onclick= function() { alert('1'); }

				if (Aglos.length > 0 ){
               if (window.showModalDialog)
               {
				hiword.onclick= function verglos2(){window.open('GLOSCONS.ASP?DATOS=' + word,window,"dialogWidth:600px;dialogHeight:350px;help:no;resizable:yes;"); return false; }
			}else
			{
				hiword.onclick= function verglos2(){window.open('GLOSCONS.ASP?DATOS=' + word,window,"Width=600,Height=350,status=no"); return false; }
				}
			}

				hiword.className = "searchword";

				hiword.appendChild(hiwordtext);



				pn.insertBefore(before,node);
				pn.insertBefore(hiword,node);

				pn.insertBefore(after,node);
				pn.removeChild(node);
			}
		}
	}

     return count_matches;
}


function highlightWordex(node,word,doc,pal) {
     doc = typeof(doc) != 'undefined' ? doc : document;
     count_matches = 0;
	// Iterate into this nodes childNodes

	//document.write(getText(node) );



	if (node.hasChildNodes) {
		var hi_cn;
		for (hi_cn=0;hi_cn<node.childNodes.length;hi_cn++) {
			count_matches = count_matches + highlightWordex(node.childNodes[hi_cn],word,doc,pal);
		}
	}

	// And do this node itself
	if (node.nodeType == 3) { // text node
		//tempNodeVal = stripVowelAccent(node.nodeValue.toLowerCase());
		//tempWordVal = stripVowelAccent(word.toLowerCase());
		//indexOf(str) === 0;
		tempNodeVal=removeDiacritics(node.nodeValue.toLowerCase());
        tempWordVal2 =word;
		tempWordVal =word;//removeDiacritics(word.toLowerCase());

      //  alert(tempNodeVal);
     // alert(tempWordVal);
     tempWordVal=tempWordVal.replace(/\s+/gi,' ')

     tempWordVal=tempWordVal.replace(' ','\\s{0,5}');

   // var number ="\\b" + tempWordVal + "\\b";

     if (tempWordVal.charAt(tempWordVal.length-1) != '.')
	         {
	         var number ="\\b" + tempWordVal + "\\b";
	 		}
	 		else
	 		{
	 		var number ="\\b" + tempWordVal ;
		}



	var tt=new RegExp(number,"gi");
	var match =tt.exec(tempNodeVal);
              if (match==null)
              {
				  ni=-1;
     		  }
			else
			{
			 ni=match.index;
             pn = node.parentNode;
			 var wlen=match[0].length;
			 pn = node.parentNode;
			 if (pn.className != "searchword") {
			 nv = node.nodeValue;
			 docWordVal=nv.substr(ni,wlen);
             if (tempWordVal.charAt(tempWordVal.length-1) == '.')
             {
             if (docWordVal.toLowerCase() != tempWordVal.toLowerCase())
									  {
										  ni=-1;

			  }

			 }}

		    }



	if (ni !== -1) {
			pn = node.parentNode;
			if (pn.className != "searchword") {
                    count_matches = count_matches + 1;
				// word has not already been highlighted!
			//	alert( node.nodeValue);
			//	alert(tempNodeVal);
				nv = node.nodeValue;
			//	ni = tempNodeVal.indexOf(tempWordVal);
				//var number ="\\b" + tempWordVal + "\\b";
				//var tt=new RegExp(number,"gi");
              //  var match =tt.exec(tempNodeVal);
              //  ni=match.index;
               // alert(nv);
				// Create a load of replacement nodes
			//	ni = tempNodeVal.indexOf(tempWordVal);
			    var wlen=match[0].length;
				before = doc.createTextNode(nv.substr(0,ni));
				docWordVal = nv.substr(ni,wlen);
				after = doc.createTextNode(nv.substr(ni+wlen));
				hiwordtext = doc.createTextNode(docWordVal);
				hiword = doc.createElement("span");
				hiword.id='sxxxs';
				//hiword.style.zIndex="999";
				//hiword.onclick= function() { alert('1'); }

				if (Aglos.length > 0 ){
                if (window.showModalDialog)
               {
				hiword.onclick= function verglos2(){window.open('GLOSCONS.ASP?DATOS=' + pal,window,"dialogWidth:600px;dialogHeight:350px;help:no;resizable:yes;"); return false;}
			}else
			{
				hiword.onclick= function verglos2(){window.open('GLOSCONS.ASP?DATOS=' + pal,window,"Width=600,Height=350,status=no"); return false; }
				}
			}

				hiword.className = "searchword";

				hiword.appendChild(hiwordtext);



				pn.insertBefore(before,node);
				pn.insertBefore(hiword,node);

				pn.insertBefore(after,node);
				pn.removeChild(node);
			}
		}
	}

     return count_matches;
}


function highlightWordok(node,word,doc) {
     doc = typeof(doc) != 'undefined' ? doc : document;
     count_matches = 0;
	// Iterate into this nodes childNodes

	//document.write(getText(node) );


	return;
	if (node.hasChildNodes) {
		var hi_cn;
		for (hi_cn=0;hi_cn<node.childNodes.length;hi_cn++) {
			count_matches = count_matches + highlightWord(node.childNodes[hi_cn],word,doc);
		}
	}

	// And do this node itself
	if (node.nodeType == 3) { // text node
		//tempNodeVal = stripVowelAccent(node.nodeValue.toLowerCase());
		//tempWordVal = stripVowelAccent(word.toLowerCase());
		//indexOf(str) === 0;
		tempNodeVal=removeDiacritics(node.nodeValue.toLowerCase());
		tempWordVal =removeDiacritics(word.toLowerCase());
    //    alert(tempNodeVal);
    //    alert(tempWordVal);
    var number ="\\b" + tempWordVal + "\\b";
	var tt=new RegExp(number,"gi");
	var match =tt.exec(tempNodeVal);
              if (match==null)
              {
				  ni=-1;
     		  }
			else
			{
			 ni=match.index;
		    }



	if (ni !== -1) {
			pn = node.parentNode;
			if (pn.className != "searchword") {
                    count_matches = count_matches + 1;
				// word has not already been highlighted!
			//	alert( node.nodeValue);
			//	alert(tempNodeVal);
				nv = node.nodeValue;
			//	ni = tempNodeVal.indexOf(tempWordVal);
				//var number ="\\b" + tempWordVal + "\\b";
				//var tt=new RegExp(number,"gi");
              //  var match =tt.exec(tempNodeVal);
              //  ni=match.index;
               // alert(nv);
				// Create a load of replacement nodes
			//	ni = tempNodeVal.indexOf(tempWordVal);
				before = doc.createTextNode(nv.substr(0,ni));
				docWordVal = nv.substr(ni,word.length);
				after = doc.createTextNode(nv.substr(ni+word.length));
				hiwordtext = doc.createTextNode(docWordVal);
				hiword = doc.createElement("span");
				hiword.id='sxxxs';
				//hiword.style.zIndex="999";
				//hiword.onclick= function() { alert('1'); }

				if (Aglos.length > 0 ){
			if (window.showModalDialog) {
				hiword.onclick= function verglos2(){window.open('GLOSCONS.ASP?DATOS=' + word,window,"dialogWidth:600px;dialogHeight:350px;help:no;resizable:yes;"); return false; }
			}else
			{
				hiword.onclick= function verglos2(){window.open('GLOSCONS.ASP?DATOS=' + word,window,"Width=600,Height=350,status=no"); return false; }
				}
			}

				hiword.className = "searchword";

				hiword.appendChild(hiwordtext);



				pn.insertBefore(before,node);
				pn.insertBefore(hiword,node);

				pn.insertBefore(after,node);
				pn.removeChild(node);
			}
		}
	}

     return count_matches;
}

function unhighlightWord(node,word,doc) {
     doc = typeof(doc) != 'undefined' ? doc : document;
	// Iterate into this nodes childNodes
	if (node.hasChildNodes) {
		var hi_cn;
		for (hi_cn=0;hi_cn<node.childNodes.length;hi_cn++) {
			unhighlightWord(node.childNodes[hi_cn],word,doc);
		}
	}

	// And do this node itself
	if (node.nodeType == 3) { // text node
		tempNodeVal = node.nodeValue.toLowerCase();
		tempWordVal = word.toLowerCase();
		if (tempNodeVal.indexOf(tempWordVal) != -1) {
			pn = node.parentNode;
			if (pn.className == "searchword") {
				prevSib = pn.previousSibling;
				nextSib = pn.nextSibling;
				nextSib.nodeValue = prevSib.nodeValue + node.nodeValue + nextSib.nodeValue;
				prevSib.nodeValue = '';
				node.nodeValue = '';
			}
		}
	}
}

function unhighlight(node) {
	// Iterate into this nodes childNodes
	if (node.hasChildNodes) {
		var hi_cn;
		for (hi_cn=0;hi_cn<node.childNodes.length;hi_cn++) {
			unhighlight(node.childNodes[hi_cn]);
		}
	}

	// And do this node itself
	if (node.nodeType == 3) { // text node
		pn = node.parentNode;
		if( pn.className == "searchword" ) {
			prevSib = pn.previousSibling;
			nextSib = pn.nextSibling;
			nextSib.nodeValue = prevSib.nodeValue + node.nodeValue + nextSib.nodeValue;
			prevSib.nodeValue = '';
			node.nodeValue = '';
		}
	}
}

function googleSearchHighlight(doc) {
     doc = typeof(doc) != 'undefined' ? doc : document;
	if (!doc.createElement) return 0;
	ref = doc.referrer;
        ref = ref.replace(/\/search\/web\//,'?search&q='); // Most WebCrawler searches
	if (ref.indexOf('?') == -1) return 0;
     count_matches = 0;
	qs = ref.substr(ref.indexOf('?')+1);
        qsa = qs.split('#');
        qs = qsa[0];
        qs = qs.replace(/(^|&)p=Q&ts=e&/,'&'); // Most Eurekster searches
        qs = qs.replace(/(^|&)query=/,'&q='); // Most Lycos searches
        qs = qs.replace(/(^|&)key=/,'&q='); // Most Walhello searches
        qs = qs.replace(/(^|&)keywords=/i,'&q='); // Most Overture searches
        qs = qs.replace(/(^|&)searchfor=/,'&q='); // Most Mysearch.com searches
        qs = qs.replace(/(^|&)qt=/,'&q='); // Most Acoona.com searches
        qs = qs.replace(/(^|&)s=/,'&q='); // Most Technirati GET searches
	qsa = qs.split('&');
	for (i=0;i<qsa.length;i++) {
		qsip = qsa[i].split('=');
	        if (qsip.length == 1) continue;
        	if (qsip[0] == 'q' || qsip[0] == 'p' || qsip[0] == 'w') { // q= for Google, p= for Yahoo, w= for Eurekster
			// Trim leading and trailing spaces after unescaping
			qsip[1] = unescape(qsip[1]).replace(/^\s+|\s+$/g, "");
			if( qsip[1] == '' ) continue;
                        phrases = qsip[1].replace(/\+/g,' ').split(/\"/);
			for(p=0;p<phrases.length;p++) {
			        phrases[p] = unescape(phrases[p]).replace(/^\s+|\s+$/g, "");
				if( phrases[p] == '' ) continue;
				if( p % 2 == 0 ) words = phrases[p].replace(/([+,()]|%(29|28)|\W+(AND|OR)\W+)/g,' ').split(/\s+/);
				else { words=Array(1); words[0] = phrases[p]; }
	                	for (w=0;w<words.length;w++) {
					if( words[w] == '' ) continue;
					count_matches = count_matches + highlightWord(doc.getElementsByTagName("body")[0],words[w],doc);
					if( p % 2 == 0 ) searchhi_string = searchhi_string + ' ' + words[w];
					else searchhi_string = searchhi_string + ' "' + words[w] + '"';
                		}
			}

	        }
	}
     return count_matches;
}

// Everything form this point on is modified to allow for highlighting
// of terms found in the REQUEST URI
function localSearchHighlight(searchStr, tryQ, doc) {
     tryQ = typeof(tryQ) != 'undefined' ? tryQ : 0;
     doc =typeof(doc) != 'undefined' ? doc : document;

	if (!doc.createElement) return 0;
     if (searchStr == '') return 0;
     count_matches = 0;
	if (searchStr.indexOf('?') == -1) qs = searchStr.substr(0);
	else qs = searchStr.substr(1);
	qsa = qs.split('&');
	for (i=0;i<qsa.length;i++) {
		qsip = qsa[i].split('=');
	        if (qsip.length == 1) continue;
        	if (qsip[0] == 'h' || ( tryQ && ( qsip[0] == 'q' || qsip[0] == 'p' ) ) ) { // be careful about ghost highlights
			// Trim leading and trailing spaces after unescaping
			qsip[1] = unescape(qsip[1]).replace(/^\s+|\s+$/g, "");
			if( qsip[1] == '' ) continue;
                        phrases = qsip[1].replace(/\+/g,' ').split(/\"/);
			// Use this next line if you would like to force the script to always
			// search for phrases. See below as well!!!
			//phrases = new Array(); phrases[0] = ''; phrases[1] = qsip[1].replace(/\+/g,' ');
			for(p=0;p<phrases.length;p++) {
			        phrases[p] = unescape(phrases[p]).replace(/^\s+|\s+$/g, "");
				if( phrases[p] == '' ) continue;

				if( p % 2 == 0 ) words = phrases[p].replace(/([+,()]|%(29|28)|\W+(AND|OR)\W+)/g,' ').split(/\s+/);
				else { words=Array(1); words[0] = phrases[p]; }
	                	for (w=0;w<words.length;w++) {
							if (words[w].length <=2 ) {continue;}
					if( words[w] == '' ) continue;
					count_matches = count_matches + highlightWord(document.getElementById('zglos'),words[w],doc);
					if( p % 2 == 0 ) searchhi_string = searchhi_string + ' ' + words[w];
					else searchhi_string = searchhi_string + ' "' + words[w] + '"';
					// As before, use this next line if forcing phrase searching
					//else searchhi_string = searchhi_string + ' ' + words[w];
                		}
			}
	        }
	}
     return count_matches;
}

function localSearchHighlight2(searchStr, tryQ, doc) {

   count_matches = 0;

  //  var div = document.createElement("div");
  //   div.style.width = "300px";
   //  div.setAttribute("id","msglos");
   //  div.style.height = "50px";
     //div.style.background = "red";
    //  div.style.position="absolute";
//
    // div.style.color = "white";
     //div.className = "searchword";
  //   div.innerHTML = "<center><span class=searchword ><br><br>    Armando Glosario.Aguarde por favor.......</span></center>";
  //   var btnas=document.getElementById('btnas');

  //   document.body.insertBefore(div);
  //  div.style.left=(document.body.clientWidth /2 )-150 + 'px';

// div.style.top=btnas.style.top;


  var elem=document.getElementById('zglos')

// document.write(getText(elem));
  var tempNodeVal=removeDiacritics(getText(elem));






  for (var w=0;w<Aglos.length;w++)
  {



	tempWordVal =removeDiacritics(Aglos[w].toLowerCase());
	//alert(tempWordVal);


	    //var number ="\\b" + tempWordVal + "\\b";
	     tempWordVal=tempWordVal.replace(/\s+/gi,' ')

        tempWordVal=tempWordVal.replace(' ','\\s{0,5}');
        if (tempWordVal.charAt(tempWordVal.length-1) != '.')
        {
        var number ="\\b" + tempWordVal + "\\b";
		}
		else
		{
		var number ="\\b" + tempWordVal ;
		}

		var tt=new RegExp(number,"gi");
		var match =tt.exec(tempNodeVal);
	              if (match==null)
	              {
					  ni=-1;
	     		  }
				else
				{
				 ni=match.index;
		    }

            if (ni==-1) {continue;}


			if (Aglos[w].length <=2 ) {continue;}
			if( Aglos[w] == '' ) continue;



				highlightWordex(elem,tempWordVal,doc,Aglos[w]);
             }

  //document.body.style.cursor = 'default	';
     return count_matches;
}

function postSearchHighlight(doc) {
        doc = typeof(doc) != 'undefined' ? doc : document;
        // Trim any leading or trailing space
        // (this is an overkill way of getting rid of the leading
        //  space that always is present in searchhi_string)
        searchhi_string = searchhi_string.replace(/^\s+|\s+$/g, "");

        // In MSIE, sometimes the dynamic generation of the spans
        // for the highlighting takes the anchor out of focus.
        // Here, we put it back in focus.
        if( doc.location.hash.length > 1 ) doc.location.hash = doc.location.hash;
}

function SearchHighlight(doc,loc) {
     // This logic should allow pages to use themselves as search
     // engines while not ghosting old searches on current results.
     doc = typeof(doc) != 'undefined' ? doc : document;
     loc = typeof(loc) != 'undefined' ? loc : doc.location;
     count_matches = 0;
     if( doc.createElement ) {
          var docrefpage = doc.referrer.split('?',2);
          var locrefpage = loc.href.split('?',1);

          // Check to see if the referrer looks like it has search terms
          //   In the special case when the referrer *IS* this page,
          //   ignore the referrer.
          if( ( docrefpage[0].toUpperCase() != locrefpage[0].toUpperCase() )
              &&
              ( ( docrefpage[1] && docrefpage[1].match(/(^|&)(p=Q&ts=e&|query=|key=|keywords=|searchfor=|qt=|q=|p=|w=|s=)/) )
                || doc.referrer.match(/\/search\/web\//) ) )
          {
               // Look to location (via "h" field) *AND* referrer
	          googleSearchHighlight(doc);
	          count_matches = count_matches + localSearchHighlight(loc.search, 0, doc);
          } else {
               // Look to location *ONLY* for highlighting
               var locsearch = loc.search;
               if( locsearch.indexOf('?') == -1 ) {
                    // Found no GET string, so look to pathname in
                    // last effort to match something
                    locsearch = locsearch.replace(/\/search\//,'?search&q='); // Most WebCrawler searches
               }
               else
               {
                    locsearch = locsearch.replace(/(\?|&)highlight=/,"$1h=");
               }
	          count_matches = count_matches + localSearchHighlight(locsearch, 1, doc);
          }
          postSearchHighlight(doc);
     }
     return count_matches;
}

function SmartHighlight(doc,loc)
{
	// This function is like SearchHighlight(doc)
	// but it detects a page refresh and toggles highlighting
	// on each refresh. This gives a quick way to turn off
	// highlighting (and quickly turn it on after).

     doc = typeof(doc) != 'undefined' ? doc : document;
     loc = typeof(loc) != 'undefined' ? loc : doc.location;

	var today = new Date();
	var now = today.getUTCSeconds();

	var cookie = doc.cookie;
	var cookieArray = cookie.split('; ');

	// Get timestamp stored in cookie
	for (var loop=0; loop < cookieArray.length; loop++){
		var nameValue = cookieArray[loop].split("=");
		if (nameValue[0].toString() == 'SHTS'){
			var cookieTime = parseInt( nameValue[1] );
		}
		else if (nameValue[0].toString() == 'SHTSP'){
			var cookieName = nameValue[1];
		}
	}

	// If we got a cookie, the cookie is from this page,
	// and the cookie's time is very close to now, then
	// this must be a page refresh (or very similar)
	// so we don't want to highlight. (the 5 second threshold
	// may need to be adjusted for slower browsers/pages/etc.)
	if( cookieName &&
		cookieTime &&
		cookieName == escape(loc.href) &&
		Math.abs(now - cookieTime) < 5 )
	{
		// Refresh detected, so don't highlight

		// Disable refresh detection for this run;
		// this is what allows us to toggle the highlighting
		// back *ON* on the next refresh
		searchhi_unl = 0;
	}
	else
	{
		// This is not a refresh, so highlight
		return SearchHighlight(doc);
	}
     return 0;

}

function SmartHLUnload(doc)
{
     doc = typeof(doc) != 'undefined' ? doc : document;
     loc = typeof(loc) != 'undefined' ? loc : doc.location;

	if( searchhi_unl > 0 )
	{
		// Turn refresh detection on so that if this
		// page gets quickly loaded, we know it's a refresh
		var today = new Date();
		var now = today.getUTCSeconds();
		doc.cookie = 'SHTS=' + now + ';';
		doc.cookie = 'SHTSP=' + escape(loc.href) + ';';
	}
	else
	{
		// Refresh detection has been disabled
		doc.cookie = 'SHTS=;';
		doc.cookie = 'SHTSP=;';
	}
}

function NotRefreshHL()
{
	// This is not a refresh. It's probably a submit
	// with the same search string, so disable refresh
	// detection on this go around.
	searchhi_unl = 0;
	return true;
}

// By default, turn refresh detection on
var searchhi_unl = 1;

// window.onload = SearchHighlight;
// window.onload = SmartHighlight;
// window.onunload = SmartHLUnload;
function ElementPosition(errorElement){
elemPosX = errorElement.offsetLeft;
    elemPosY = errorElement.offsetTop;
    theElement = errorElement.offsetParent;
    alert(elemPosX);
    alert(elemPosY);
    alert(theElement);
    while(theElement != null){
      elemPosY += theElement.offsetTop;
      theElement = theElement.offsetParent;
    }
    window.scrollTo(0,elemPosY);
}