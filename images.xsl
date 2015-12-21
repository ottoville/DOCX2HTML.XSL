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
		<xsl:param name="themefile" />
		<xsl:param name="size" />
		<xsl:variable name="dravingid" select="pic:blipFill/a:blip/@r:embed"/>
		<img>
			<xsl:attribute name="style">
				<xsl:choose>
					<xsl:when test="local-name(../../..)='anchor'">
						float:left;position:absolute;
						top:<xsl:apply-templates select="../../../wp:positionV/wp:posOffset"/>;
						left:<xsl:apply-templates select="../../../wp:positionH/wp:posOffset"/>;
						z-index:<xsl:value-of select="../../../@relativeHeight"/>;
					</xsl:when>
					<xsl:otherwise>
						display:inline;
					</xsl:otherwise>
				</xsl:choose>
			width:<xsl:value-of select="number($size/@cx) div 360000"/>cm;height:<xsl:value-of select="number($size/@cy) div 360000"/>cm</xsl:attribute>
			<xsl:attribute name="src"><xsl:value-of select="resolve-uri(document($reldocument)/*/*[@Id=$dravingid]/@Target,base-uri())"/></xsl:attribute>
		</img>
	</xsl:template>
	<xsl:template match="mc:AlternateContent">
		<xsl:param name="reldocument" />
		<xsl:param name="themefile" />
		<xsl:apply-templates select="mc:Choice">
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="themefile" select="$themefile" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="mc:Choice">
		<xsl:param name="reldocument" />
		<xsl:param name="themefile" />
		<xsl:apply-templates select="w:drawing">
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="themefile" select="$themefile" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="wp:posOffset">
		<xsl:value-of select="((number(text()) div 91440)*8)*(4 div 3)"/>px
	</xsl:template>
	<xsl:template match="wps:wsp">
		<xsl:param name="reldocument" />
		<xsl:param name="themefile" />
		<xsl:param name="size" />
		<svg:svg>
			<xsl:attribute name="style">
				<xsl:choose>
					<xsl:when test="local-name(../../..)='anchor'">
						float:left;position:absolute;
						top:<xsl:apply-templates select="../../../wp:positionV/wp:posOffset"/>;
						left:<xsl:apply-templates select="../../../wp:positionH/wp:posOffset"/>;
						z-index:<xsl:value-of select="../../../@relativeHeight"/>;
					</xsl:when>
					<xsl:otherwise>
						display:inline;
					</xsl:otherwise>
				</xsl:choose>
				<xsl:if test="wps:spPr/a:xfrm/@rot">
					transform:rotate(<xsl:value-of select="number(wps:spPr/a:xfrm/@rot) div 60000"/>deg);
				</xsl:if>
				width:<xsl:value-of select="((number($size/@cx) div 91440) * 8) *(4 div 3)"/>px;height:<xsl:value-of select="((number($size/@cy) div 91440) * 8) *(4 div 3)"/>px;
			</xsl:attribute>
			<xsl:apply-templates select="wps:style">
				<xsl:with-param name="reldocument" select="$reldocument" />
				<xsl:with-param name="themefile" select="$themefile" />
			</xsl:apply-templates>
			<xsl:apply-templates select="wps:spPr">
				<xsl:with-param name="reldocument" select="$reldocument" />
				<xsl:with-param name="themefile" select="$themefile" />
			</xsl:apply-templates>
		</svg:svg>
	</xsl:template>
	<xsl:template match="wps:style">
		<xsl:param name="themefile" />
		<xsl:attribute name="class">
			<xsl:if test="a:fillRef">
				<xsl:value-of select="concat(a:fillRef/a:schemeClr/@val,'fill ')"/>
			</xsl:if>
			<xsl:if test="a:lnRef">
				<xsl:value-of select="concat(a:lnRef/a:schemeClr/@val,'border ')"/>
			</xsl:if>
		</xsl:attribute>
	</xsl:template>
	<xsl:template match="a:schemeClr">
		<xsl:param name="reldocument" />
		<xsl:param name="themefile" />
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="themefile" select="$themefile" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="wps:spPr">
		<xsl:param name="reldocument" />
		<xsl:param name="themefile" />
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="themefile" select="$themefile" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template name="svgclasses">
		<xsl:param name="themefile" />
		<xsl:for-each select="../*">
			<xsl:choose>
				<xsl:when test="local-name(.)='solidFill'">
					<xsl:value-of select="concat(a:schemeClr/@val,' ')" />
				</xsl:when>
				<!--<xsl:when test="local-name(.)='ln'">
					<xsl:value-of select="concat(a:solidFill/a:schemeClr/@val,'Border ')" />
				</xsl:when>-->
			</xsl:choose>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="svgElement">
		<xsl:param name="themefile" />
		<xsl:attribute name="class">
			<xsl:call-template name="svgclasses">
				<xsl:with-param name="themefile" select="$themefile" />
			</xsl:call-template>
		</xsl:attribute>
		<xsl:attribute name="style">
			<xsl:call-template name="svgstyles">
				<xsl:with-param name="themefile" select="$themefile" />
			</xsl:call-template>
		</xsl:attribute>
	</xsl:template>
	<xsl:template name="hextodec">
		<xsl:param name="hex"/>
		<xsl:choose>
			<xsl:when test="$hex='A' or $hex='a'">
				<xsl:value-of select="10"/>
			</xsl:when>	
			<xsl:when test="$hex='B' or $hex='b'">
				<xsl:value-of select="11"/>
			</xsl:when>
			<xsl:when test="$hex='C' or $hex='c'">
				<xsl:value-of select="12"/>
			</xsl:when>
			<xsl:when test="$hex='D' or $hex='d'">
				<xsl:value-of select="13"/>
			</xsl:when>
			<xsl:when test="$hex='E' or $hex='e'">
				<xsl:value-of select="14"/>
			</xsl:when>
			<xsl:when test="$hex='F' or $hex='f'">
				<xsl:value-of select="15"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="number($hex)"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="rgbtohsl">
		<xsl:param name="r"/>
		<xsl:param name="g"/>
		<xsl:param name="b"/>
		<xsl:param name="modify"/>
		<xsl:variable name="r2" select="$r div 255" />
		<xsl:variable name="g2" select="$g div 255" />
		<xsl:variable name="b2" select="$b div 255" />
		<xsl:variable name="max" select="max(($r2,$g2,$b2))" />
		<xsl:variable name="min" select="min(($r2,$g2,$b2))" />
		<xsl:variable name="d" select="$max - $min" />
		<xsl:variable name="l" select="($max + $min) div 2" />
		<xsl:variable name="lumoff" select="100000 - number($modify)" />
		<xsl:variable name="l2">
			<xsl:choose>
				<xsl:when test="number($modify) &gt; 0">
					<xsl:value-of select="$l* number($modify) div 100000"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="($max + $min) div 2"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="s">
			<xsl:choose>
				<xsl:when test="$l=0">
					<xsl:value-of select="0"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="$d div (1 - (2 * $l - 1))"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="h">
			<xsl:choose>
				<xsl:when test="$max=$r2">
					<xsl:choose>
						<xsl:when test="$g2 &lt; $b2">
							<xsl:value-of select="(($g2 - $b2) div $d + 6) div 6"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="(($g2 - $b2) div $d) div 6"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$max=$g2">
					<xsl:value-of select="(($b2 - $r2) div $d + 2) div 6"/>
				</xsl:when>
				<xsl:when test="$max=$b2">
					<xsl:value-of select="(($r2 - $g2) div $d + 4) div 6"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<xsl:value-of select="concat($h * 360,',',$s * 100,'%,',$l2 * 100,'%')" />
	</xsl:template>
	<xsl:template name="hexToDecimal">
		<xsl:param name="hex"/>
		<xsl:variable name="firstdecnum">
			<xsl:call-template name="hextodec"><xsl:with-param name="hex" select="substring($hex,1,1)"/></xsl:call-template>
		</xsl:variable>
		<xsl:variable name="seconddecnum">
			<xsl:call-template name="hextodec"><xsl:with-param name="hex" select="substring($hex,2,1)"/></xsl:call-template>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="string-length($hex) &gt; 1">
				<xsl:value-of select="($firstdecnum * 16) + $seconddecnum"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$firstdecnum"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="fillcolor" match="a:solidFill">
		<xsl:param name="themefile" />
		<xsl:variable name="schemeClr" select="a:schemeClr/@val"/>
		<xsl:variable name="color">
			<xsl:choose>
				<xsl:when test="a:schemeClr">
					<xsl:value-of select="$themefile/a:theme/a:themeElements/a:clrScheme/a:*[local-name(.)=$schemeClr]/a:srgbClr/@val"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="a:srgbClr/@val"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		hsl(<xsl:call-template name="rgbtohsl">
			<xsl:with-param name="r">
				<xsl:call-template name="hexToDecimal"><xsl:with-param name="hex" select="substring($color,1,2)"/></xsl:call-template>
			</xsl:with-param>
			<xsl:with-param name="g">
				<xsl:call-template name="hexToDecimal"><xsl:with-param name="hex" select="substring($color,3,2)"/></xsl:call-template>
			</xsl:with-param>
			<xsl:with-param name="b">
				<xsl:call-template name="hexToDecimal"><xsl:with-param name="hex" select="substring($color,5,2)"/></xsl:call-template>
			</xsl:with-param>
			<xsl:with-param name="modify">
				<xsl:choose>
					<xsl:when test="a:schemeClr">
						<xsl:value-of select="a:schemeClr/a:lumMod/@val"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="a:srgbClr/a:lumMod/@val"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:with-param>
		</xsl:call-template>);
	</xsl:template>
	<xsl:template name="svgstyles">
		<xsl:param name="themefile" />
		<xsl:for-each select="../*">
			<xsl:choose>
				<xsl:when test="local-name(.)='solidFill'">
					fill:<xsl:apply-templates select=".">
						<xsl:with-param name="themefile" select="$themefile" />
					</xsl:apply-templates>
				</xsl:when>
				<xsl:when test="local-name(.)='ln'">
					stroke:<xsl:apply-templates select="a:solidFill">
						<xsl:with-param name="themefile" select="$themefile" />
					</xsl:apply-templates>
					stroke-width:2;
				</xsl:when>
			</xsl:choose>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="a:prstGeom[@prst='ellipse']">
		<xsl:param name="themefile" />
		<xsl:variable name="x" select="((number(../a:xfrm/a:ext/@cx) div 91440) * 8) *(4 div 3) div 2" />
		<xsl:variable name="y" select="((number(../a:xfrm/a:ext/@cy) div 91440) * 8) *(4 div 3) div 2" />
		<svg:ellipse cy="{$y + number(../a:xfrm/a:off/@y)}" cx="{$x + number(../a:xfrm/a:off/@x)}" rx="{$x}" ry="{$y}">
			<xsl:call-template name="svgElement">
				<xsl:with-param name="themefile" select="$themefile" />
			</xsl:call-template>
		</svg:ellipse>
	</xsl:template>
	<xsl:template match="a:prstGeom[@prst='corner']">
		<xsl:param name="themefile" />
		<xsl:variable name="x" select="((number(../a:xfrm/a:ext/@cx) div 91440) * 8) *(4 div 3)" />
		<xsl:variable name="y" select="((number(../a:xfrm/a:ext/@cy) div 91440) * 8) *(4 div 3)" />
		<xsl:variable name="p1" select="concat(0,',',0)" />
		<xsl:variable name="p2" select="concat($x div 2,',',0)" />
		<xsl:variable name="p3" select="concat($x div 2,',',$y div 2)" />
		<xsl:variable name="p4" select="concat($x,',',$y div 2)" />
		<xsl:variable name="p5" select="concat($x,',',$y)" />
		<xsl:variable name="p6" select="concat(0,',',$y)" />
		<!--
		0,0
		0.5,0
		0.5,0.5
		1,0.5
		1,1
		0,1
		-->
		<svg:polygon points="{concat($p1,' ',$p2,' ',$p3,' ', $p4,' ',$p5,' ',$p6)}">
			<xsl:call-template name="svgElement">
				<xsl:with-param name="themefile" select="$themefile" />
			</xsl:call-template>
		</svg:polygon>
	</xsl:template>
	<xsl:template match="a:prstGeom[@prst='triangle']">
		<xsl:param name="themefile" />
		<xsl:variable name="x" select="((number(../a:xfrm/a:ext/@cx) div 91440) * 8) *(4 div 3)" />
		<xsl:variable name="y" select="((number(../a:xfrm/a:ext/@cy) div 91440) * 8) *(4 div 3)" />
		<xsl:variable name="p1" select="concat($x div 2,',',0)" />
		<xsl:variable name="p2" select="concat($x,',',$y)" />
		<xsl:variable name="p3" select="concat(0,',',$y)" />
		<!--
		0.5,0	1,1		0,1
		-->
		<svg:polygon points="{concat($p1,' ',$p2,' ',$p3)}">
			<xsl:call-template name="svgElement">
				<xsl:with-param name="themefile" select="$themefile" />
			</xsl:call-template>
		</svg:polygon>
	</xsl:template>
	<xsl:template match="a:prstGeom[@prst='hexagon']">
		<xsl:param name="themefile" />
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
			<xsl:call-template name="svgElement">
				<xsl:with-param name="themefile" select="$themefile" />
			</xsl:call-template>
		</svg:polygon>
	</xsl:template>
	<xsl:template match="wp:anchor">
		<xsl:param name="reldocument" />
		<xsl:param name="themefile" />
		<xsl:apply-templates select="a:graphic">
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="size" select="wp:extent" />
			<xsl:with-param name="themefile" select="$themefile" />
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
		<xsl:param name="themefile" />		
		<xsl:param name="size" />		
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="size" select="$size" />
			<xsl:with-param name="themefile" select="$themefile" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="a:graphicData">
		<xsl:param name="reldocument" />		
		<xsl:param name="themefile" />		
		<xsl:param name="size" />		
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="themefile" select="$themefile" />
			<xsl:with-param name="size" select="$size" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="w:drawing">
		<xsl:param name="reldocument" />
		<xsl:param name="themefile" />
		<xsl:apply-templates select="wp:anchor[wp:positionH/@relativeFrom!='page' and wp:positionV/@relativeFrom!='page']|wp:inline">
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="themefile" select="$themefile" />
		</xsl:apply-templates>
	</xsl:template>
</xsl:stylesheet>