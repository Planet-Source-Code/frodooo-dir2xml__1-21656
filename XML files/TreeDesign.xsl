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
