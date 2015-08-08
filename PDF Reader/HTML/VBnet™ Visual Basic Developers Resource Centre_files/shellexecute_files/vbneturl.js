{ 
var path = "";
var href = document.location.href;
var url= "";
var s = href.split("/"); 

for (var i=3;i<(s.length-1);i++) {
   path+="<A HREF=\""+href.substring(0,href.indexOf(s[i])+s[i].length)+"/\">"+s[i]+"</A> / ";
}

i=s.length-1;
path+="<A HREF=\""+href.substring(0,href.indexOf(s[i])+s[i].length)+"\">"+s[i]+"</A>";

url = '<font color="#29527C" size="1" face="verdana, arial">';
url = url + "Page: " + path ;
url = url + '</font>';

document.writeln(url);
}
