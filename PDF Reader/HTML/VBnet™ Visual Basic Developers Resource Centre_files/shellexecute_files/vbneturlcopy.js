function copyURLtoClipboard()
   { 
     window.clipboardData.clearData;
     window.clipboardData.setData("Text",document.location.href);
   }

function copySearchURLtoClipboard()
   { 
     window.clipboardData.clearData;
     window.clipboardData.setData("Text","http://vbnet.mvps.org/search/main/index.html");
   }

function postShortcut1(v) 
{
  window.clipboardData.clearData;
  document.all.holdtext.value = v;
  document.all.holdtext.createTextRange().execCommand("Copy");
}

function postShortcut2() 
{
  window.clipboardData.clearData;
  document.all.holdtext.value = document.location.href;
  document.all.holdtext.createTextRange().execCommand("Copy");
}

function postShortcut3(v) 
{
  window.clipboardData.clearData;
  document.all.holdtext.value = v;
  document.all.holdtext.createTextRange().execCommand("Copy");
}
