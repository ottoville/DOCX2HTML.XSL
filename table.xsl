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
	<xsl:template match="w:tc">
		<xsl:param name="reldocument" />
		<xsl:if test="count(w:tcPr/w:vMerge[not(@w:val='restart')])=0">
			<td>
				<xsl:if test="count(w:tcPr/w:vMerge[@w:val='restart'])">
					<xsl:variable name="rowposition" select="count(../preceding-sibling::w:tr) + 1" />
					<xsl:variable name="curpos" select="position()" />
					<xsl:variable name="nextrestart">
						<xsl:choose>
							<xsl:when test="count(../following-sibling::w:tr/w:tc/w:tcPr/w:vMerge[@w:val='restart']/../../../preceding-sibling::w:tr)">
								<xsl:value-of select="count(../following-sibling::w:tr/w:tc/w:tcPr/w:vMerge[@w:val='restart']/../../../preceding-sibling::w:tr)+1"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'no'"/>
							</xsl:otherwise>
						</xsl:choose>				
					</xsl:variable>
					<xsl:choose>
						<xsl:when test="$nextrestart='no'">
							<xsl:attribute name="rowspan">
								<xsl:value-of select="count(../following-sibling::w:tr/w:tc/w:tcPr/w:vMerge)+1"/>
							</xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:for-each select="../following-sibling::w:tr[position() &lt; $nextrestart]/*[position()=$curpos]/w:tcPr/w:vMerge">
								<xsl:variable name="thisrowposition" select="count(./preceding-sibling::w:tr) + 1 - $rowposition" />

								<xsl:value-of select="$thisrowposition"/>|
							</xsl:for-each>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
				<xsl:if test="count(w:tcPr/w:gridSpan)">
					<xsl:attribute name="colspan"><xsl:value-of select="w:tcPr/w:gridSpan/@w:val"/></xsl:attribute>
				</xsl:if>
				<xsl:apply-templates select="w:tcPr|w:p|w:sdt">
					<xsl:with-param name="reldocument" select="$reldocument" />
				</xsl:apply-templates>
			</td>
		</xsl:if>
	</xsl:template>
	<xsl:template match="w:tr">
		<xsl:param name="reldocument" />
		<xsl:param name="trscopeselector">tr[data-rid="<xsl:value-of select="generate-id(.)"/>"]</xsl:param>
		<tr>
			<xsl:attribute name="data-rid" select="generate-id(.)"/>
			<xsl:apply-templates select="w:trPr|w:tc|w:sdt">
				<xsl:with-param name="scopeselector" select="$trscopeselector" />
				<xsl:with-param name="reldocument" select="$reldocument" />
			</xsl:apply-templates>
		</tr>
	</xsl:template>
	<xsl:template match="w:tbl">
		<xsl:param name="reldocument" />
		<xsl:param name="scopeselector">table[data-tableid="<xsl:number/>"]</xsl:param>
		<table>
			<xsl:attribute name="data-tableid"><xsl:number/></xsl:attribute>
			<xsl:apply-templates select="w:tblPr|w:tblGrid|w:tr">
				<xsl:with-param name="reldocument" select="$reldocument" />
				<xsl:with-param name="scopeselector" select="$scopeselector" />
			</xsl:apply-templates>
		</table>
	</xsl:template>
</xsl:stylesheet>