
// swapMenu script by Randy Birch - vbnet.mvps.org/
// update 2000.09.02
// update 2000.09.23

function swapMenu(navName, OnOff) 
{
  var menumenupath;
  var ext;
  var obj;
  var suffix ;
  menupath = "../../images/nav/";
  ext = ".gif";
  suffix = ''
  if (OnOff == 1)
     { suffix = '_on' 
     }

  if (document.images);  
    obj = eval("document.images." + navName) ;
    { 
       if (obj != null);     
       { 
          obj.src = (menupath + navName + suffix  + ext);
       }
    }
}

