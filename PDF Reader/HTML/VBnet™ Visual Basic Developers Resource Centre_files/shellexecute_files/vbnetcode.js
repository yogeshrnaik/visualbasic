// Navigation script by Karl E. Peterson - www.mvps.org/vb/
// FrontPage check script by Randy Birch - vbnet.mvps.org/

function isFrontPageDesign() { 
//   alert(window.location.href);
   if (window.location.href.indexOf("Temporary%20Internet%20Files") != -1) 
          return true ;
     else return false ;
}

if (isFrontPageDesign() == false)  {

   if (top == self || (parent.frames[1].name != 'text')) {

      var thisPage = window.location.href;
      var relUrl = thisPage.substring((thisPage.indexOf('code')));
      var newURL = '../../index.html?' + relUrl;
      if (document.images)
         top.location.replace(newURL);
      else
         top.location.href = newURL;
   }
}


function isMsie4orGreater() { 
  var ua = window.navigator.userAgent;
  var msie = ua.indexOf ( "MSIE " );
  
  if  (msie > 0)
    {return (parseInt ( ua.substring ( msie+5, ua.indexOf ( ".", msie ) ) ) >=4) && (ua.indexOf("MSIE 4.0b") < 0 ) ;}
  else {return false;}
}
