Dir2XML ReadMe

by Frodooo (Frodooo@Hotmail.com), March 2001.



This Dir2XML converts a directory structure to an XML file. Just like the Dir2HTML application I posted earlier, 
this XML-exporter can be used in a browser as a sitemap for example. The easy part was to export the directory structure from VB in an XML 
file since this is more or less the same as for the HTML export. The difficult part was to write down a DTD to make it a valid document, 
an XSL transform to render it, some JavaScript for DHTML and a CSS fro the design. Altough the DTD is not strictly necessary. 
You can simply drop the DTD-string (sDTD) in the VB code without any problems, but it's a good test. 
The XLS transform is tricky but fun anyway. You can use it as I did or use it to transform it via ASP on the server side, 
which is quite easy. In fact you need something like


<%
	sXml = "/IDesk/Admin/Tree/Explorer.xml"
	sXsl ="/IDesk/Admin/Tree/TreeDesign.xsl"	
	set oXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM")
	set oXslDoc = Server.CreateObject("MICROSOFT.XMLDOM")
	oXmlDoc.async = false
	oXslDoc.async = false
	oXmlDoc.load(Server.MapPath(sXml))
	If oXmlDoc.parseError.errorcode <> 0 Then
   		Response.Write oXmlDoc.parseError.errorCode	
	End If
	oXslDoc.load(Server.MapPath(sXsl))
	If oXslDoc.parseError.errorcode <> 0 Then
   		Response.Write oXslDoc.parseError.errorCode	
	End If
	Response.Write oXmlDoc.transformNode(oXslDoc)	
%>

It would have been nice to parametrize the export just like the HTML exporter by allowing the user to specify the location of the JavaScript, the XSL and the CSS. Unfortunately this is not so easy since these files are closely related with each other. As you can see below e.g. the XSL contains the location of the JavaScript and as such it is not straightforward to change it via the VB interface. I did not say it's not possible but it requires a bit of time.

For those who do not understand the logic behind the app let me summarize it:

- the user has chosen a directory
- this directory is used as the root of the export
- VB starts by writing the mendatory XML lines 
- a recursive function is called  (GoThroughDir) which contains the main code to export the directory structure to XML. Note that you can modify the attributes of the XML but be careful, the XSL depends on them and case-sensitivity is around
- in the XML there's a line which refers to the XSL, when the browser opens the XML file the XML is converted to HTML
- in the XSL a CSS is specified and a link to some JavaScript
- when the browser iterates through the XML nodes it checks the XSL templates and renders the HTML accordingly
- some nodes have a CSS class attribute, this is the way to change the appearance of the nodes
- the JavaScripts comes in when the user clicks or moves the mouse over the HTML


Hope this helps, if you need assistance just send me a note. Have fun!



_____________________________________________________________________________________________________________________________
The XML Style Transformation (TreeDesign.xsl)





<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
	<!-- start tree template -->
	<xsl:template match="/">
		
		<!-- import script and CSS  -->
		<LINK REL="stylesheet" TYPE="text/css" HREF="TreeDesign.css"/>
		<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="TreeScripts.js"></SCRIPT>		
	
		<xsl:apply-templates select="folders"/>
				<br/><br/>
				<A class="clsButton" href="#+" onclick="ShowAll('UL')">Expand all</A>   <A class="clsButton" href="#-" onclick="HideAll('UL')">Collapse all</A>
				<br/><br/>

	<UL>
		<xsl:apply-templates select="folders/folder"/>
		<xsl:apply-templates select="folders/file"/>

	</UL>
	</xsl:template>
	<!-- end template -->
	<xsl:template match="folders">
			<SPAN class="clsTitle">
			<b>
				<xsl:value-of select="@DIRNAME"/>
			</b>
			</SPAN>
		
	</xsl:template>
	<xsl:template match="file">
	<LI>
		<A TARGET="Main">
		<xsl:attribute name="HREF">
			<xsl:value-of select="URL"/>
		</xsl:attribute>
		<xsl:value-of select="TITLE"/>
		</A>
	</LI>
	</xsl:template>
	<xsl:template match="folder">
	
		<LI CLASS="clsHasKids">
			<SPAN>
				<xsl:value-of select="@DIRNAME"/>
			</SPAN>
			<UL>
				<xsl:for-each select="file">
					<LI>
						<A TARGET="Main">
							<xsl:attribute name="HREF">
								<xsl:value-of select="URL"/>
							</xsl:attribute>
							<xsl:value-of select="TITLE"/>
						</A>
					</LI>
				</xsl:for-each>
				<xsl:apply-templates select="folder"/>
			</UL>
		</LI>
	</xsl:template>
</xsl:stylesheet>



_____________________________________________________________________________________________________________________________

Document Type Definition (DTD)

<!DOCTYPE folders[ 
<!ELEMENT folders (folder|file)+>
<!ELEMENT folder (file| folder)*>
<!ELEMENT file (TITLE, URL)>
<!ELEMENT TITLE (#PCDATA)>
<!ELEMENT URL (#PCDATA)>
<!ATTLIST folders
DIRNAME CDATA #REQUIRED
ID ID #REQUIRED
>
<!ATTLIST folder
DIRNAME CDATA #REQUIRED
ID ID #REQUIRED
>
<!ATTLIST file
FILENAME CDATA #REQUIRED
ID ID #REQUIRED
>
]>

_____________________________________________________________________________________________________________________________


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
_____________________________________________________________________________________________________________________________




	/* TreeDesign.css */
	
	BODY { font-family:verdana; font-size:70%; }
	H1 { font-size:120%; font-style:italic; }

	UL { margin-left:0px; margin-bottom:5px; }
	LI UL { display:none; margin-left:16px; }
	LI { font-weight:bold; list-style-type:square; cursor:default; text-indent:10px;}
	LI.clsHasKids { list-style-type:none;  }
	LI.clsHasKids SPAN { text-Indent:10pt ; cursor:hand; font-weight:bold; font-family:verdana; font-size:110%; list-style-	image:URL(images/FClosed.gif) }

	A:link, A:visited, A:active { font-weight:normal; color:navy; }
	A:hover { text-decoration:none; }

	BUTTON { font-family:tahoma; font-size:100%; }
	.clsTitle {background-color:steelblue;color:white;}
	.clsButton {background-color:steelblue;color:white;}
	
	
	
_____________________________________________________________________________________________________________________________	