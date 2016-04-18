<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xhtml="http://www.w3.org/1999/xhtml"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
	xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
	xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
	xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
	xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
	xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
	xmlns:svg="http://www.w3.org/2000/svg"
    exclude-result-prefixes="xs w r pr wp a pic xhtml w14 wps mc svg"
    version="2.0">
	<!-- docx2html.xsl
		Project started by Otto-Ville Lamminpää
			ottoville.lamminpaa@gmail.com
			+358445596869
	-->	
	<xsl:template name="page">
		<xsl:param name="reldocument" />
		<xsl:param name="sectionselector" />
		<xsl:param name="themefile" />
		<xsl:param name="selector" />
		<xsl:param name="defaultheader" />
		<xsl:param name="defaultfooter" />
		<xsl:param name="evenheader" />
		<xsl:param name="evenfooter" />
		<xsl:param name="firstpageheader" />
		<xsl:param name="firstpagefooter" />
		<xsl:param name="sectionnumber" />
		<!--<xsl:if test="$selector/w:r|$selector/w:tr">--> <!-- Filter empty pages where text is marked to continue on next page but there are no more text -->
			<div>
				<xsl:attribute name="class" select="$sectionnumber" />
				<xsl:if test="string-length($firstpageheader)">
					<header>
						<xsl:attribute name="class">first</xsl:attribute>
						<xsl:apply-templates select="document(resolve-uri(document($reldocument)/*/*[@Id=$firstpageheader]/@Target,base-uri()))/*/*">
							<xsl:with-param name="reldocument" select="resolve-uri(concat('_rels/',document($reldocument)/*/*[@Id=$firstpageheader]/@Target,'.rels'),base-uri())" />
							<xsl:with-param name="themefile" select="$themefile" />
						</xsl:apply-templates>
					</header>
				</xsl:if>
				<xsl:if test="string-length($evenheader)">
					<header>
						<xsl:attribute name="class">even</xsl:attribute>
						<xsl:apply-templates select="document(resolve-uri(document($reldocument)/*/*[@Id=$evenheader]/@Target,base-uri()))/*/*">
							<xsl:with-param name="reldocument" select="resolve-uri(concat('_rels/',document($reldocument)/*/*[@Id=$evenheader]/@Target,'.rels'),base-uri())" />
							<xsl:with-param name="themefile" select="$themefile" />
						</xsl:apply-templates>
					</header>
				</xsl:if>
				<xsl:if test="string-length($defaultheader)">
					<header>
						<xsl:attribute name="class">default</xsl:attribute>
						<xsl:apply-templates select="document(resolve-uri(document($reldocument)/*/*[@Id=$defaultheader]/@Target,base-uri()))/*/*">
							<xsl:with-param name="reldocument" select="resolve-uri(concat('_rels/',document($reldocument)/*/*[@Id=$defaultheader]/@Target,'.rels'),base-uri())" />
							<xsl:with-param name="themefile" select="$themefile" />
						</xsl:apply-templates>
					</header>
				</xsl:if>		
				<xsl:for-each select="$selector">
					<xsl:variable name="listid" select="w:pPr/w:numPr/w:numId/@w:val" />
					<xsl:variable name="nextlistid" select="./following-sibling::w:p[w:pPr/w:numPr][not(w:pPr/w:numPr/w:numId/@w:val=$listid)][1]/w:pPr/w:numPr/w:numId/@w:val" />
					<!--
					Filter empty paragraphs?
					<xsl:if test="not(./preceding-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=$nextlistid]) and (not(w:pPr/w:numPr) or not(./preceding-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=current()/w:pPr/w:numPr/w:numId/@w:val]))">
						-->
					<xsl:apply-templates select="current()">
							<xsl:with-param name="reldocument" select="$reldocument" />
							<xsl:with-param name="themefile" select="$themefile" />
						</xsl:apply-templates>
					<!--</xsl:if> -->
				</xsl:for-each>	
				<xsl:if test="string-length($firstpagefooter)">
					<footer>
						<xsl:attribute name="class">first</xsl:attribute>
						<xsl:apply-templates select="document(resolve-uri(document($reldocument)/*/*[@Id=$firstpagefooter]/@Target,base-uri()))/*/*">
							<xsl:with-param name="reldocument" select="resolve-uri(concat('_rels/',document($reldocument)/*/*[@Id=$firstpagefooter]/@Target,'.rels'),base-uri())" />
							<xsl:with-param name="themefile" select="$themefile" />
						</xsl:apply-templates>
					</footer>
				</xsl:if>
				<xsl:if test="string-length($evenfooter)">
					<footer>
						<xsl:attribute name="class">even</xsl:attribute>
						<xsl:apply-templates select="document(resolve-uri(document($reldocument)/*/*[@Id=$evenfooter]/@Target,base-uri()))/*/*">
							<xsl:with-param name="reldocument" select="resolve-uri(concat('_rels/',document($reldocument)/*/*[@Id=$evenfooter]/@Target,'.rels'),base-uri())" />
							<xsl:with-param name="themefile" select="$themefile" />
						</xsl:apply-templates>
					</footer>
				</xsl:if>
				<xsl:if test="string-length($defaultfooter)">
					<footer>
						<xsl:attribute name="class">default</xsl:attribute>
						<xsl:apply-templates select="document(resolve-uri(document($reldocument)/*/*[@Id=$defaultfooter]/@Target,base-uri()))/*/*">
							<xsl:with-param name="reldocument" select="resolve-uri(concat('_rels/',document($reldocument)/*/*[@Id=$defaultfooter]/@Target,'.rels'),base-uri())" />
							<xsl:with-param name="themefile" select="$themefile" />
						</xsl:apply-templates>
					</footer>
				</xsl:if>				
			</div>
		<!--</xsl:if>-->
	</xsl:template>
	<xsl:template name="pages"> <!-- always last element of page -->
		<xsl:param name="reldocument" />
		<xsl:param name="themefile" />
		<xsl:param name="prevpage" />
		<xsl:param name="selector" />
		<xsl:param name="sectionselector" />
		<xsl:param name="defaultheader" />
		<xsl:param name="defaultfooter" />
		<xsl:param name="firstpageheader" />
		<xsl:param name="firstpagefooter" />
		<xsl:param name="evenheader" />
		<xsl:param name="evenfooter" />
		<xsl:param name="sectionnumber" />
		<xsl:choose>
			<xsl:when test="not($prevpage)">
				<xsl:call-template name="page">
					<xsl:with-param name="reldocument" select="$reldocument" />
					<xsl:with-param name="sectionselector" select="$sectionselector" />
					<xsl:with-param name="sectionnumber" select="$sectionnumber" />
					<xsl:with-param name="themefile" select="$themefile" />
					<xsl:with-param name="defaultheader" select="$defaultheader" />
					<xsl:with-param name="defaultfooter" select="$defaultfooter" />
					<xsl:with-param name="firstpageheader" select="$firstpageheader" />
					<xsl:with-param name="firstpagefooter" select="$firstpagefooter" />
					<xsl:with-param name="evenheader" select="$evenheader" />
					<xsl:with-param name="selector" select="$selector[count(preceding-sibling::*) &lt; count(current()/preceding-sibling::*)+1]" />
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:call-template name="page">
					<xsl:with-param name="sectionnumber" select="$sectionnumber" />
					<xsl:with-param name="reldocument" select="$reldocument" />
					<xsl:with-param name="sectionselector" select="$sectionselector" />
					<xsl:with-param name="themefile" select="$themefile" />
					<xsl:with-param name="defaultheader" select="$defaultheader" />
					<xsl:with-param name="defaultfooter" select="$defaultfooter" />
					<xsl:with-param name="firstpageheader" select="$firstpageheader" />
					<xsl:with-param name="firstpagefooter" select="$firstpagefooter" />
					<xsl:with-param name="evenheader" select="$evenheader" />
					<xsl:with-param name="evenfooter" select="$evenfooter" />
					<xsl:with-param name="selector" select="$prevpage/following-sibling::*[count(./preceding-sibling::*) &lt; count(current()/preceding-sibling::*)+1]" />
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="sections"><!-- Always the last element of section -->
		<xsl:param name="themefile" />
		<xsl:variable name="reldocument" select="resolve-uri('_rels/document.xml.rels',base-uri())" />

		<!-- Either the page section, following section or document section-->
		<xsl:variable name="sectionselector" select="w:pPr/w:sectPr|./following-sibling::*[w:pPr/w:sectPr][1]/w:pPr/w:sectPr|/w:document/w:body/w:sectPr"/>
		<xsl:variable name="thisid" select="generate-id(current())" />
		<xsl:variable name="paddingtop" select="number($sectionselector[1]/w:pgMar/@w:top)" />
		<xsl:variable name="paddingbottom" select="number($sectionselector[1]/w:pgMar/@w:bottom)" />
		<xsl:variable name="paddingright" select="number($sectionselector[1]/w:pgMar/@w:right)" />
		<xsl:variable name="paddingleft" select="number($sectionselector[1]/w:pgMar/@w:left)" />
		<xsl:variable name="prevsection" select="preceding-sibling::*[w:pPr/w:sectPr][1]" />		
		<xsl:variable name="sectionnumber" select="concat('section',position())" />
		<xsl:variable name="defaultheader">
			<xsl:choose>
				<xsl:when test="$sectionselector/w:headerReference[@w:type='default']/@r:id">
					<xsl:value-of select="$sectionselector/w:headerReference[@w:type='default']/@r:id" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="preceding-sibling::*[w:pPr/w:sectPr/w:headerReference[@w:type='default']][1]/w:pPr/w:sectPr[1]/w:headerReference[@w:type='default']/@r:id" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="defaultfooter">
			<xsl:choose>
				<xsl:when test="$sectionselector/w:footerReference[@w:type='default']/@r:id">
					<xsl:value-of select="$sectionselector/w:footerReference[@w:type='default']/@r:id" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="preceding-sibling::*[w:pPr/w:sectPr/w:footerReference[@w:type='default']][1]/w:pPr/w:sectPr[1]/w:footerReference[@w:type='default']/@r:id" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="evenheader">
			<xsl:choose>
				<xsl:when test="$sectionselector/w:headerReference[@w:type='even']/@r:id">
					<xsl:value-of select="$sectionselector/w:headerReference[@w:type='even']/@r:id" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="preceding-sibling::*[w:pPr/w:sectPr/w:headerReference[@w:type='even']][1]/w:pPr/w:sectPr[1]/w:headerReference[@w:type='even']/@r:id" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="evenfooter">
			<xsl:choose>
				<xsl:when test="$sectionselector/w:footerReference[@w:type='even']/@r:id">
					<xsl:value-of select="$sectionselector/w:footerReference[@w:type='even']/@r:id" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="preceding-sibling::*[w:pPr/w:sectPr/w:footerReference[@w:type='even']][1]/w:pPr/w:sectPr[1]/w:footerReference[@w:type='even']/@r:id" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="firstpageheader">
			<xsl:choose>
				<xsl:when test="$sectionselector/w:headerReference[@w:type='first']/@r:id">
					<xsl:value-of select="$sectionselector/w:headerReference[@w:type='first']/@r:id" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="preceding-sibling::*[w:pPr/w:sectPr/w:headerReference[@w:type='first']][1]/w:pPr/w:sectPr[1]/w:headerReference[@w:type='first']/@r:id" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="firstpagefooter">
			<xsl:choose>
				<xsl:when test="$sectionselector/w:footerReference[@w:type='first']/@r:id">
					<xsl:value-of select="$sectionselector/w:footerReference[@w:type='first']/@r:id" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="preceding-sibling::*[w:pPr/w:sectPr/w:footerReference[@w:type='first']][1]/w:pPr/w:sectPr[1]/w:footerReference[@w:type='first']/@r:id" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<hr>
			<xsl:attribute name="class" select="$sectionnumber" />
			<xsl:if test="$sectionselector/w:titlePg">
				<xsl:attribute name="data-firstheader">true</xsl:attribute>
			</xsl:if>
		</hr>
		<!-- Apply template to pages that belongs to this section -->
		<xsl:choose>
			<xsl:when test="not($prevsection)"> <!--first section or document with one section -->
			<xsl:for-each select="(preceding-sibling::*|current())[w:r/w:br[@w:type='page']]|(preceding-sibling::*|current())[last()]"> <!-- always last element of page -->
					<xsl:call-template name="pages">
						<xsl:with-param name="reldocument" select="$reldocument" />
						<xsl:with-param name="themefile" select="$themefile" />
						<xsl:with-param name="sectionselector" select="$sectionselector" />
						<xsl:with-param name="selector" select="preceding-sibling::*|current()" />
						<xsl:with-param name="defaultheader" select="$defaultheader" />
						<xsl:with-param name="defaultfooter" select="$defaultfooter" />
						<xsl:with-param name="evenheader" select="$evenheader" />
						<xsl:with-param name="evenfooter" select="$evenfooter" />
						<xsl:with-param name="firstpageheader" select="$firstpageheader" />
						<xsl:with-param name="firstpagefooter" select="$firstpagefooter" />
						<xsl:with-param name="prevpage" select="(preceding-sibling::*|current())[w:r/w:br[@w:type='page']][count(./preceding-sibling::*) &lt; count(current()/preceding-sibling::*)][last()]" />
						<xsl:with-param name="sectionnumber" select="$sectionnumber" />
					</xsl:call-template>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test="w:pPr/w:sectPr">
				<xsl:for-each select="(current()|./preceding-sibling::*[count(./preceding-sibling::*) &gt; count($prevsection/preceding-sibling::*)])[w:r/w:br[@w:type='page']]|(current()|./preceding-sibling::*[count(./preceding-sibling::*) &gt; count($prevsection/preceding-sibling::*)])[last()]">
					<xsl:call-template name="pages">
						<xsl:with-param name="reldocument" select="$reldocument" />
						<xsl:with-param name="themefile" select="$themefile" />
						<xsl:with-param name="sectionselector" select="$sectionselector" />
						<xsl:with-param name="selector" select="current()|./preceding-sibling::*[count(./preceding-sibling::*) &gt; count($prevsection/preceding-sibling::*)]" />
						<xsl:with-param name="prevpage" select="(current()|./preceding-sibling::*[count(./preceding-sibling::*) &gt; count($prevsection/preceding-sibling::*)])[w:r/w:br[@w:type='page']][count(./preceding-sibling::*) &lt; count(current()/preceding-sibling::*)][last()]" />
						<xsl:with-param name="defaultheader" select="$defaultheader" />
						<xsl:with-param name="defaultfooter" select="$defaultfooter" />
						<xsl:with-param name="evenheader" select="$evenheader" />
						<xsl:with-param name="evenfooter" select="$evenfooter" />
						<xsl:with-param name="firstpageheader" select="$firstpageheader" />
						<xsl:with-param name="firstpagefooter" select="$firstpagefooter" />
						<xsl:with-param name="sectionnumber" select="$sectionnumber" />
					</xsl:call-template>
				</xsl:for-each>
			</xsl:when>
			<xsl:otherwise> <!-- Last section -->
				<xsl:for-each select="($prevsection/following-sibling::*)[w:r/w:br[@w:type='page']]|($prevsection/following-sibling::*)[last()]">
					<xsl:call-template name="pages">
						<xsl:with-param name="reldocument" select="$reldocument" />
						<xsl:with-param name="themefile" select="$themefile" />
						<xsl:with-param name="sectionselector" select="$sectionselector" />
						<xsl:with-param name="selector" select="$prevsection/following-sibling::*" />
						<xsl:with-param name="prevpage" select="($prevsection/following-sibling::*)[w:r/w:br[@w:type='page']][count(./preceding-sibling::*) &lt; count(current()/preceding-sibling::*)][last()]" />
						<xsl:with-param name="defaultheader" select="$defaultheader" />
						<xsl:with-param name="defaultfooter" select="$defaultfooter" />
						<xsl:with-param name="evenheader" select="$evenheader" />
						<xsl:with-param name="evenfooter" select="$evenfooter" />
						<xsl:with-param name="firstpageheader" select="$firstpageheader" />
						<xsl:with-param name="firstpagefooter" select="$firstpagefooter" />
						<xsl:with-param name="sectionnumber" select="$sectionnumber" />
					</xsl:call-template>
				</xsl:for-each>
			</xsl:otherwise>
		</xsl:choose>
		<style>
			<xsl:value-of select="concat('div.',$sectionnumber,' {')" />
				<xsl:value-of select="concat('width:',(number($sectionselector[1]/w:pgSz/@w:w) - $paddingright - $paddingleft) div 20,'pt;')"/>
				<xsl:value-of select="concat('min-height:',(number($sectionselector[1]/w:pgSz/@w:h) - $paddingtop - $paddingbottom) div 20,'pt;')"/>
				<xsl:value-of select="concat('padding:',$paddingtop div 20,'pt ',$paddingright div 20,'pt ',$paddingbottom div 20,'pt ',$paddingleft div 20,'pt;')"/>
			 }
		</style>
	</xsl:template>
</xsl:stylesheet>