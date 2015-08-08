var d1='mvps'
var d2='org'
var addr='rgb'
function GoEmail(title, doc)
{
   var stmp='mailto:'+addr+'@'+d1+'.'+d2+'?Subject='+title+'&Body=re: '+doc+'%0D%0A'+'%0D%0A';
   window.location.replace(stmp);
}


function GoMainEmail()
{
   stmp='mailto:'+addr+'@'+d1+'.'+d2+'?Subject='+parent.frames[1].document.title+'&Body=re: '+parent.frames[1].document.location+'%0D%0A'+'%0D%0A';
   window.location.replace(stmp);
}

function GoMainEmail2(title, doc)
{
   stmp='mailto:'+addr+'@'+d1+'.'+d2+'?Subject='+title+'&Body=re: '+doc+'%0D%0A'+'%0D%0A';
   window.location.replace(stmp);
}


function trimString(val) //string trim function 
{ 
  var valu = val; 
  valu = rtrimString(valu); //trim leading spaces 
  valu = ltrimString(valu); //trim trailing spaces 
  return valu; //return trimed variable 
} 

function rtrimString(val) //remove trailing spaces 
{ 
  var valu = val; 
  while (valu.charAt(valu.length - 1) == " ") //trim all trailing spaces 
  { 
    valu = valu.substr(0,valu.length - 1); //trim a trailing space 
  } 
  return valu; 
} 

function ltrimString(val) //remove leading spaces 
{ 
  var valu = val; 
  while (valu.charAt(0) == " ") //trim all leading spaces 
  { 
    valu = valu.substr(1); //trim leading space 
  } 
  return valu; 
} 

function GoMainEmailWithTrimmedTitle(title, doc)
{
   var revTitle= ltrimString(title.substring((title.indexOf(']'))+1));
   stmp='mailto:'+addr+'@'+d1+'.'+d2+'?Subject='+revTitle+'&Body=re: '+doc+'%0D%0A'+'%0D%0A';
   window.location.replace(stmp);    

}


