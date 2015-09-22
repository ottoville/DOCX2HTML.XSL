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
	<xsl:template match="pic:pic">
		<xsl:param name="reldocument" />
		<xsl:param name="size" />
		<xsl:variable name="dravingid" select="pic:blipFill/a:blip/@r:embed"/>
		<img>
			<xsl:attribute name="style">display:inline;width:<xsl:value-of select="number($size/@cx) div 360000"/>cm;height:<xsl:value-of select="number($size/@cy) div 360000"/>cm</xsl:attribute>
			<xsl:attribute name="src"><xsl:value-of select="resolve-uri(document($reldocument)/*/*[@Id=$dravingid]/@Target,base-uri())"/></xsl:attribute>
		</img>
	</xsl:template>
	<xsl:template match="mc:AlternateContent">
		<xsl:param name="reldocument" />
		<xsl:apply-templates select="mc:Choice">
			<xsl:with-param name="reldocument" select="$reldocument" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="mc:Choice">
		<xsl:param name="reldocument" />
		<xsl:apply-templates select="w:drawing">
			<xsl:with-param name="reldocument" select="$reldocument" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="wps:wsp">
		<xsl:param name="reldocument" />
		<xsl:param name="size" />
		<svg:svg>
			<xsl:attribute name="style">
				width:<xsl:value-of select="((number($size/@cx) div 91440) * 8) *(4 div 3)"/>px;height:<xsl:value-of select="((number($size/@cy) div 91440) * 8) *(4 div 3)"/>px;
			</xsl:attribute>
			<xsl:apply-templates>
				<xsl:with-param name="reldocument" select="$reldocument" />
			</xsl:apply-templates>
		</svg:svg>
	</xsl:template>
	<xsl:template match="wps:style">
		<xsl:param name="reldocument" />
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="a:schemeClr">
		<xsl:param name="reldocument" />
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="wps:spPr">
		<xsl:param name="reldocument" />
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="a:prstGeom[@prst='ellipse']">
		<xsl:variable name="x" select="((number(../a:xfrm/a:ext/@cx) div 91440) * 8) *(4 div 3) div 2" />
		<xsl:variable name="y" select="((number(../a:xfrm/a:ext/@cy) div 91440) * 8) *(4 div 3) div 2" />
		<svg:ellipse cy="{$y + number(../a:xfrm/a:off/@y)}" cx="{$x + number(../a:xfrm/a:off/@x)}" rx="{$x}" ry="{$y}">
			<xsl:attribute name="class" select="../../wps:style/a:fillRef/a:schemeClr/@val"/>
		</svg:ellipse>
	</xsl:template>
	<xsl:template match="a:prstGeom[@prst='hexagon']">
		<xsl:variable name="x" select="((number(../a:xfrm/a:ext/@cx) div 91440) * 8) *(4 div 3)" />
		<xsl:variable name="y" select="((number(../a:xfrm/a:ext/@cy) div 91440) * 8) *(4 div 3)" />
		<xsl:variable name="p1" select="concat($x div 4,',',0)" />
		<xsl:variable name="p2" select="concat($x div 4 * 3,',',0)" />
		<xsl:variable name="p3" select="concat($x,',',$y div 2)" />
		<xsl:variable name="p4" select="concat($x div 4 *3,',',$y)" />
		<xsl:variable name="p5" select="concat($x div 4,',',$y)" />
		<xsl:variable name="p6" select="concat(0,',',$y div 2)" />
		<!--
		<xsl:variable name="p1" select="'1,0'" />
		<xsl:variable name="p2" select="'3,0'" />
		<xsl:variable name="p3" select="'4,2'" />
		<xsl:variable name="p4" select="'3,4'" />
		<xsl:variable name="p5" select="'1,4'" />
		<xsl:variable name="p6" select="'0,2'" />
		-->
		<svg:polygon points="{concat($p1,' ',$p2,' ',$p3,' ', $p4,' ',$p5,' ',$p6)}">
			<xsl:attribute name="class" select="../../wps:style/a:fillRef/a:schemeClr/@val"/>
		</svg:polygon>
	</xsl:template>
	<xsl:template match="wp:anchor">
		<xsl:param name="reldocument" />
		<xsl:apply-templates select="a:graphic">
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="size" select="wp:extent" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="wp:inline">
		<xsl:param name="reldocument" />		
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="size" select="wp:extent" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="a:graphic">
		<xsl:param name="reldocument" />		
		<xsl:param name="size" />		
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="size" select="$size" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="a:graphicData">
		<xsl:param name="reldocument" />		
		<xsl:param name="size" />		
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="size" select="$size" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="w:drawing">
		<xsl:param name="reldocument" />
		<xsl:apply-templates select="wp:anchor|wp:inline">
			<xsl:with-param name="reldocument" select="$reldocument" />
		</xsl:apply-templates>
	</xsl:template>
</xsl:stylesheet>