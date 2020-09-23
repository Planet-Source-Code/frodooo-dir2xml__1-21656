
	/* TreeScripts.js */
	
  function GetChildElem(eSrc,sTagName)
  {
    var cKids = eSrc.children;
    for (var i=0;i<cKids.length;i++)
    {
      if (sTagName == cKids[i].tagName) return cKids[i];
    }
    return false;
  }
  
  function document.onclick()
  {
    var eSrc = window.event.srcElement;
		if ("SPAN" == eSrc.tagName && "clsHasKids" == eSrc.parentElement.className)
		{var eChild = GetChildElem(eSrc.parentElement,"UL");
      		eChild.style.display = ("block" == eChild.style.display ? "none" : "block");      		
      		if (eChild.style.display=="block")
      			{eSrc.style.listStyleImage="URL('images/FOpen.gif')"}
      		else
      			{eSrc.style.listStyleImage="URL('images/FClosed.gif')"};
    }
  }

  function document.onmouseover()
  {
    var eSrc = window.event.srcElement;
		if ("SPAN" == eSrc.tagName && "clsHasKids" == eSrc.parentElement.className)
		{
			eSrc.style.color = "maroon";
    		};
    		if ("A" == eSrc.tagName && "clsButton" == eSrc.parentElement.className)
		{
			eSrc.style.color = "maroon";
    		}
  }

  function document.onmouseout()
  {
    var eSrc = window.event.srcElement;
		if ("SPAN" == eSrc.tagName && "clsHasKids" == eSrc.parentElement.className)
		{
			eSrc.style.color = "";
    		};
    		if ("A" == eSrc.tagName && "clsButton" == eSrc.parentElement.className)
		{
			eSrc.style.color = "";
    		}
  }

  function ShowAll(sTagName)
  {
    var cElems = document.all.tags(sTagName);
    var iNumElems = cElems.length;
    for (var i=1;i<iNumElems;i++) cElems[i].style.display = "block";
  }
  
  function HideAll(sTagName)
  {
    var cElems = document.all.tags(sTagName);
    var iNumElems = cElems.length;
    for (var i=1;i<iNumElems;i++) cElems[i].style.display = "none";
  }
