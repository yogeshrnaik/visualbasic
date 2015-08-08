// IE code-clip by Randy Birch - vbnet.mvps.org/

function ClipBas() 
   {
      holdtext.innerText = getCopyright() + copybas.innerText + '\n\n';
      Copied = holdtext.createTextRange();
      Copied.execCommand("RemoveFormat");      
      Copied.execCommand("Copy");
   }

function ClipBas2() 
   {
      holdtext.innerText = getCopyright() + copybas2.innerText + '\n\n';
      Copied = holdtext.createTextRange();
      Copied.execCommand("RemoveFormat");      
      Copied.execCommand("Copy");
   }

function ClipBas3() 
   {
      holdtext.innerText = getCopyright() + copybas3.innerText + '\n\n';
      Copied = holdtext.createTextRange();
      Copied.execCommand("RemoveFormat");      
      Copied.execCommand("Copy");
   }


function ClipForm() 
   {
      holdtext.innerText = getCopyright() + copyfrm.innerText+ '\n\n';
      Copied = holdtext.createTextRange();
      Copied.execCommand("RemoveFormat");      
      Copied.execCommand("Copy");
   }

function ClipForm2() 
   {
      holdtext.innerText = getCopyright() + copyfrm2.innerText+ '\n\n';
      Copied = holdtext.createTextRange();
      Copied.execCommand("RemoveFormat");      
      Copied.execCommand("Copy");
   }

function ClipForm3() 
   {
      holdtext.innerText = getCopyright() + copyfrm3.innerText+ '\n\n';
      Copied = holdtext.createTextRange();
      Copied.execCommand("RemoveFormat");      
      Copied.execCommand("Copy");
   }

function ClipClass() 
   {
      holdtext.innerText = getCopyright() + copyclass.innerText+ '\n\n';
      Copied = holdtext.createTextRange();
      Copied.execCommand("RemoveFormat");      
      Copied.execCommand("Copy");
   }


function ClipSnipPrivate() 
   {
      holdtext.innerText = getCopyright() + copyall.innerText+ '\n\n';
      Copied = holdtext.createTextRange();
      Copied.execCommand("RemoveFormat");      
      Copied.execCommand("Copy");
   }

function ClipSnipPublic() 
   {
      holdtext.innerText = getCopyright() + copyall.innerText+ '\n\n';
      holdtext.innerText = replaceChars(holdtext.innerText);
      Copied = holdtext.createTextRange();
      Copied.execCommand("RemoveFormat");
      
      Copied.execCommand("Copy");
   }

function ClipSnipDeclare() 
   {
      holdtext.innerText = getCopyright() + copydeclare.innerText+ '\n\n';
      Copied = holdtext.createTextRange();
      Copied.execCommand("RemoveFormat");      
      Copied.execCommand("Copy");
   }

function getCopyright() 
   {
	var credit1 = "'-----------------------------------------------------------------------------------------\n";
	var credit2 = "' Copyright ©1996-2006 VBnet, Randy Birch. All Rights Reserved Worldwide.\n";
	var credit3 = "'        Terms of use http://vbnet.mvps.org/terms/pages/terms.htm\n";
	var credit4 = "'-----------------------------------------------------------------------------------------\n\n";
	return (credit1 + credit2 + credit3 + credit4);	
   }

function replaceChars(entry) {
out = "Private"; // replace this
add = "Public"; // with this
temp = "" + entry; // temporary holder

while (temp.indexOf(out)>-1) {
pos= temp.indexOf(out);
temp = "" + (temp.substring(0, pos) + add + 
temp.substring((pos + out.length), temp.length));
}
return temp;
}

