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
    exclude-result-prefixes="xs w r pr wp a pic xhtml w14 wps"
    version="2.0">
<!-- List continue-->
	<xsl:template name="listitem" match="w:p[w:pPr/w:numPr and count(./preceding-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=current()/w:pPr/w:numPr/w:numId/@w:val]) &gt; 0 ]">
		<xsl:param name="reldocument" />
		<xsl:param name="listintend"/>
		<xsl:param name="thismargin">
			<xsl:choose>
				<xsl:when test="w:pPr/w:ind/@w:left">
					<xsl:value-of select="number(w:pPr/w:ind/@w:left)"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="number(0)"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:param>
		<xsl:variable name="listid" select="w:pPr/w:numPr/w:numId/@w:val" />
		<xsl:variable name="currentid" select="generate-id(current())" />
		<xsl:variable name="nextlistid" select="following-sibling::w:p[w:pPr/w:numPr][not(w:pPr/w:numPr/w:numId/@w:val=$listid)][1]/w:pPr/w:numPr/w:numId/@w:val" />
		<xsl:variable name="prevlistid" select="preceding-sibling::w:p[w:pPr/w:numPr][not(w:pPr/w:numPr/w:numId/@w:val=$listid)][1]/w:pPr/w:numPr/w:numId/@w:val" />
		<li>
			<xsl:attribute name="style">
				<xsl:if test="number($thismargin)>0">
					margin-left:<xsl:value-of select='(number($thismargin)-$listintend) div 20'/>pt;
				</xsl:if>
				<xsl:if test="w:pPr/w:ind/@w:hanging">
					text-indent:-<xsl:value-of select='number(w:pPr/w:ind/@w:hanging) div 20'/>pt;
				</xsl:if>
			</xsl:attribute>
			<xsl:attribute name="id"><xsl:value-of select="generate-id(current())"/></xsl:attribute>
			<xsl:call-template name="paragraph">
				<xsl:with-param name="reldocument" select="$reldocument" />
				<xsl:with-param name="listintend" select="$listintend + $thismargin" />
			</xsl:call-template>
			<xsl:apply-templates select="./following-sibling::w:p[
														generate-id(./preceding-sibling::w:p[
															w:pPr/w:numPr
														][1])=$currentid
													 and (not(w:pPr/w:numPr/w:numId/@w:val=$listid) and not(w:pPr/w:numPr/w:numId/@w:val=$prevlistid))
													 and (./following-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=$listid] or
														count(./preceding-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=$nextlistid])  &gt; 0)]">
				<xsl:with-param name="reldocument" select="$reldocument" />
				<xsl:with-param name="listintend" select="$listintend + $thismargin" />
			</xsl:apply-templates>
		</li>
	</xsl:template>
	<xsl:template name="paragraph" match="w:p[not(w:pPr/w:numPr)]">
		<xsl:param name="reldocument" />
		<xsl:param name="listintend"/>
		<xsl:variable name="currentid" select="generate-id(current())" />
		<xsl:variable name="class">
			<xsl:if test="count(w:pPr/w:pStyle)">
				<xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
			</xsl:if>
			<xsl:if test="local-name(..)='sdtContent'">
				<xsl:value-of select="concat(' ',generate-id(../..))"/>
			</xsl:if>
		</xsl:variable>
		<p class="{normalize-space($class)}" pid="{$currentid}">
			<xsl:apply-templates select="w:pPr|w:r[
										(w:fldChar[@w:fldCharType='begin'] or 
										(not(w:fldChar|w:instrText) and not(./preceding-sibling::w:r[w:fldChar][1]/w:fldChar/@w:fldCharType='separate')))
										and 
											not(local-name(./preceding-sibling::*[1])='r' and deep-equal(./w:rPr|./w:tab|./w:fldChar, ./preceding-sibling::*[1]/w:rPr|./preceding-sibling::*[1]/w:tab|./preceding-sibling::*[1]/w:fldChar))
										]|w:ins|w:hyperlink|w:sdt">
				<xsl:with-param name="scopeselector">p[pid='<xsl:value-of select="$currentid"/>']</xsl:with-param>
				<xsl:with-param name="listintend" select="$listintend" />
				<xsl:with-param name="reldocument" select="$reldocument" />
			</xsl:apply-templates>
			<xsl:if test='count(w:r)=0'>
				<br />
			</xsl:if>
		</p>
	</xsl:template>
	<!-- List begin-->
	<xsl:template match="w:p[w:pPr/w:numPr and not(./preceding-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=current()/w:pPr/w:numPr/w:numId/@w:val])]">
		<xsl:param name="reldocument" />
		<xsl:param name="listid" select="w:pPr/w:numPr/w:numId/@w:val" />
		<xsl:param name="listlevel" select="w:pPr/w:numPr/w:ilvl/@w:val" />
		<xsl:param name="relid" select="document(resolve-uri('numbering.xml',base-uri()))/w:numbering/w:num[@w:numId=$listid]/w:abstractNumId/@w:val" />
		<xsl:param name="thismargin">
			<xsl:choose>
				<xsl:when test="w:pPr/w:ind/@w:left">
					<xsl:value-of select="number(w:pPr/w:ind/@w:left)"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="number(0)"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:param>
		<xsl:param name="thislistintend">
			<xsl:choose>
				<xsl:when test="document(resolve-uri('numbering.xml',base-uri()))/w:numbering/w:abstractNum[@w:abstractNumId=$relid]/w:lvl[@w:ilvl=$listlevel]/w:pPr/w:ind/@w:left"><xsl:value-of select="number(document(resolve-uri('numbering.xml',base-uri()))/w:numbering/w:abstractNum[@w:abstractNumId=$relid]/w:lvl[@w:ilvl=$listlevel]/w:pPr/w:ind/@w:left)"/></xsl:when>
				<xsl:otherwise><xsl:value-of select="number(0)"/></xsl:otherwise>
			</xsl:choose>
		</xsl:param>
		<xsl:variable name="elemid" select="generate-id(current())" />
		<ul listmargin="{$thismargin}" listinend="{$thislistintend}">
			<xsl:attribute name="class">level<xsl:value-of select="w:pPr/w:numPr/w:ilvl/@w:val"/> list<xsl:value-of select="w:pPr/w:numPr/w:numId/@w:val"/></xsl:attribute>
			<xsl:call-template name="listitem">
				<xsl:with-param name="listintend" select="number($thislistintend)" />
			</xsl:call-template>
			<xsl:apply-templates select="./following-sibling::w:p[(
																		w:pPr/w:numPr and
																		w:pPr/w:numPr/w:numId/@w:val=$listid)]">
				<xsl:with-param name="reldocument" select="$reldocument" />
				<xsl:with-param name="listintend" select="number($thislistintend)" />
			</xsl:apply-templates>
		</ul>
	</xsl:template>
</xsl:stylesheet>