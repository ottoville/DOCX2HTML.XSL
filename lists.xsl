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
    exclude-result-prefixes="xs w r pr wp a pic xhtml w14 wps mc"
    version="2.0">
	<xsl:template match="w:num">
		<xsl:variable name="listid" select="@w:numId"/>
		<xsl:variable name="absid" select="w:abstractNumId/@w:val"/>
		<xsl:for-each select="../w:abstractNum[@w:abstractNumId=$absid]/w:lvl[w:pPr|w:rPr]">
			<xsl:value-of select="concat('ul.level',@w:ilvl,'.list',$listid,' {')"/> 
				<xsl:if test="w:pPr/w:ind/@w:left">
					<xsl:value-of select="concat('padding-left:',number(w:pPr/w:ind/@w:left) div 20,'pt;')"/>
				</xsl:if>
			}
			<xsl:value-of select="concat('ul.level',@w:ilvl,'.list',$listid,'&gt;li {')"/>
				<xsl:if test="w:pPr/w:ind/@w:hanging">
					<xsl:value-of select="concat('text-indent:-',number(w:pPr/w:ind/@w:hanging) div 20,'pt;')"/>
				</xsl:if>
			}
			<xsl:value-of select="concat('ul.level',@w:ilvl,'.list',$listid,'>&gt;li&gt;p:first-child {')"/>
				<xsl:apply-templates select="w:pPr" /> 
			}
			<xsl:value-of select="concat('ul.level',@w:ilvl,'.list',$listid,'&gt;li&gt;p:first-child&gt;span { ')"/><xsl:apply-templates select="w:rPr" /> }
		</xsl:for-each>
	</xsl:template>
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
</xsl:stylesheet>
