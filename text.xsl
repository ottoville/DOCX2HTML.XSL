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
	<xsl:template match="w:rPr" name="spanstyle">
		<xsl:if test="../w:t/@xml:space='preserve'">white-space:pre-wrap;</xsl:if>
		<xsl:for-each select="./*">
			<xsl:choose>
				<xsl:when test="local-name(.)='spacing'">letter-spacing:<xsl:value-of select='number(@w:val) div 12'/>pt;</xsl:when>
				<xsl:when test='local-name(.)="kern" and number(../w:sz/@w:val) &gt;= number(@w:val)'>
					font-kerning:auto;
				</xsl:when>
				<xsl:when test="local-name(.)='shd'">background-color:#<xsl:value-of select="@w:fill"/>;</xsl:when>
				<xsl:when test="local-name(.)='rFonts'">font-family:<xsl:value-of select="@w:*[1]"/>;</xsl:when>
				<xsl:when test="local-name(.)='color'">color:#<xsl:value-of select="@w:val"/>;</xsl:when>
				<xsl:when test="local-name(.)='sz' or local-name(.)='szCs'">
					<xsl:choose>
						<xsl:when test='number(@w:val) &gt; 0'>
							font-size:<xsl:value-of select='number(@w:val) div 2'/>pt;
						</xsl:when>
						<xsl:otherwise>
							font-size:<xsl:value-of select="0"/>;
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="local-name(.)='b'">font-weight:bold;</xsl:when>
				<xsl:when test="local-name(.)='i'">font-style:italic;</xsl:when>
				<xsl:when test="local-name(.)='u'">text-decoration:underline;</xsl:when>
			</xsl:choose>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="w:t">
		<xsl:value-of select="."/>
	</xsl:template>
	<xsl:template match="w:r[w:tab]">
		<xsl:param name="reldocument" />
		<xsl:variable name="prevtab" select="preceding-sibling::*[w:tab][1]" />
		<xsl:variable name="tabstop" select="(number(document(resolve-uri('settings.xml',base-uri()))/w:settings/w:defaultTabStop/@w:val) div 20) * (4 div 3) "/>
		<span>
			<xsl:attribute name="class">tab</xsl:attribute>
			<xsl:attribute name="style">
				display:inline-block;
				min-width:<xsl:value-of select="$tabstop"/>px;
				text-indent:0;
			</xsl:attribute>
			<xsl:choose>
				<xsl:when test="not($prevtab)">
					<xsl:apply-templates select="./preceding-sibling::w:r[(w:fldChar[@w:fldCharType='begin'] or 
										(not(w:fldChar|w:instrText) and not(./preceding-sibling::w:r[w:fldChar][1]/w:fldChar/@w:fldCharType='separate')))]">
						<xsl:with-param name="reldocument" select="$reldocument" />
					</xsl:apply-templates>
				</xsl:when>
				<xsl:otherwise>
					<xsl:for-each select="$prevtab">
						<xsl:call-template name="text">
							<xsl:with-param name="reldocument" select="$reldocument" />
						</xsl:call-template>
					</xsl:for-each>
					<xsl:apply-templates select="./preceding-sibling::*[(count(./preceding-sibling::*) &gt; count($prevtab/preceding-sibling::*)) and (w:fldChar[@w:fldCharType='begin'] or 
										(not(w:fldChar|w:instrText) and not(./preceding-sibling::w:r[w:fldChar][1]/w:fldChar/@w:fldCharType='separate')))]">
						<xsl:with-param name="reldocument" select="$reldocument" />
					</xsl:apply-templates>
				</xsl:otherwise>
			</xsl:choose>
		</span>
		<xsl:if test="../w:pPr/w:ind/@w:hanging and not(./preceding-sibling::w:r[w:tab])">
			<br />
		</xsl:if>
		<xsl:if test="not(./following-sibling::w:r[w:tab])">
			<xsl:call-template name="text">
				<xsl:with-param name="reldocument" select="$reldocument" />
			</xsl:call-template>			
		</xsl:if>
	</xsl:template>
	<xsl:template name="fieldtext" match="w:r[./preceding-sibling::w:r[w:fldChar][1]/w:fldChar/@w:fldCharType='separate']">
	</xsl:template>
	<xsl:template name="text" match="w:r[not(w:tab) and not(./preceding-sibling::w:r[w:fldChar][1]/w:fldChar/@w:fldCharType='separate')]">
		<xsl:param name="reldocument" />
		<xsl:variable name="class">
			<xsl:if test="count(w:rPr/w:rStyle)">
				<xsl:value-of select="concat(' ',w:rPr/w:rStyle/@w:val)"/>
			</xsl:if>
			<xsl:if test="local-name(..)='sdtContent'">
				<xsl:value-of select="concat(' ',generate-id(../..))"/>
			</xsl:if>
			<xsl:if test="count(../../w:pPr/w:rPr/w:rStyle)">
				<xsl:value-of select="concat(' ',../../w:pPr/w:rPr/w:rStyle/@w:val)"/>
			</xsl:if>
		</xsl:variable>
		<span>
			<xsl:attribute name="class" select="normalize-space($class)"/>
			<xsl:attribute name="style">
				<xsl:apply-templates select="w:rPr"/>
			</xsl:attribute>
			<xsl:apply-templates select="w:t|w:drawing|mc:AlternateContent">
				<xsl:with-param name="reldocument" select="$reldocument" />
			</xsl:apply-templates>
		</span>
	</xsl:template>
	<xsl:template match="w:r[w:fldChar[@w:fldCharType='begin']]">
		<xsl:param name="reldocument" />
		<xsl:variable name="fieldtype">
			<xsl:choose>
				<xsl:when test="contains(./following-sibling::w:r[1]/w:instrText,'&#34;')">
					<xsl:value-of select="normalize-space(substring-before(./following-sibling::w:r[1]/w:instrText,'&#34;'))"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="normalize-space(./following-sibling::w:r[1]/w:instrText)"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="end" select="./following-sibling::w:r[w:fldChar/@w:fldCharType='end'][1]" />
		<xsl:variable name="startid" select="generate-id(./following-sibling::w:r[w:fldChar/@w:fldCharType='separate'][1])" />
		<xsl:variable name="text" select="$end/preceding-sibling::w:r[generate-id(./preceding-sibling::w:r[w:fldChar/@w:fldCharType='separate'][1])=$startid]" />
		<xsl:variable name="type">
			<xsl:choose>
				<xsl:when test="$fieldtype='FORMCHECKBOX'">
					<xsl:value-of select="'checkbox'" />
				</xsl:when>
				<xsl:when test="$fieldtype='HYPERLINK'">
					<xsl:value-of select="'hyperlink'" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="'text'" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$type='hyperlink'">
				<a href="{substring-before(substring-after(./following-sibling::w:r[1]/w:instrText,'&#34;'),'&#34;')}">
					<xsl:attribute name="style">
							<xsl:apply-templates select="w:rPr"/>
					</xsl:attribute>
					<xsl:for-each select="$text">
						<xsl:call-template name="text">
							<xsl:with-param name="reldocument" select="$reldocument" />
						</xsl:call-template>
					</xsl:for-each>
				</a>
			</xsl:when>
			<xsl:otherwise>
				<input type="{$type}" placeholder="{$text}"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>