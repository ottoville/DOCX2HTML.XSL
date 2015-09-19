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
	<!-- docx2html.xsl
		Project started by Otto-Ville Lamminpää
			ottoville.lamminpaa@gmail.com
			+358445596869
	-->
	<xsl:import href="paragraphs.xsl"/>
	<xsl:import href="text.xsl"/>
	<xsl:import href="table.xsl"/>
	<xsl:template name="borders">
		<xsl:for-each select="*">
			<xsl:variable name="border-style">
				<xsl:choose>
					<xsl:when test="@w:val='single' or @w:val='thick'">
						<xsl:value-of select="'solid'"/>
					</xsl:when>
					<xsl:when test="@w:val='doubleWave'">
						<xsl:value-of select="'groove'"/>
					</xsl:when>
					<xsl:when test="@w:val='nil'">
						<xsl:value-of select="'none'"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="@w:val"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:variable>
			<xsl:choose>
				<xsl:when test="$border-style='none'">
					border-<xsl:value-of select="local-name(.)"/>-style:none;
				</xsl:when>
				<xsl:otherwise>
					<xsl:choose>
						<xsl:when test="@w:color='auto'">
							border-<xsl:value-of select="local-name(.)"/>-color:inherit;
						</xsl:when>
						<xsl:otherwise>
							border-<xsl:value-of select="local-name(.)"/>-color:#<xsl:value-of select='@w:color'/>;
						</xsl:otherwise>
					</xsl:choose>
					border-<xsl:value-of select="local-name(.)"/>: <xsl:value-of select='number(@w:sz) div 4'/>px <xsl:value-of select='$border-style'/>;
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="w:ind">
		<xsl:param name="listintend"/>
		<xsl:for-each select="./@*">
			<xsl:choose>
				<xsl:when test="local-name(.)='hanging'">
					text-indent:-<xsl:value-of select='number(.) div 20'/>pt;
				</xsl:when>
				<xsl:when test="local-name(.)='firstLine'">
					text-indent:<xsl:value-of select='number(.) div 20'/>pt;
				</xsl:when>
				<xsl:when test="local-name(.)='left'">
					<xsl:choose>
						<xsl:when test="number($listintend) &gt; 0">
							padding-left:<xsl:value-of select='(number(.)-$listintend) div 20'/>pt;orgpadding:<xsl:value-of select='number(.)'/>;listindend:<xsl:value-of select='$listintend'/>;
						</xsl:when>
						<xsl:otherwise>
							padding-left:<xsl:value-of select='number(.) div 20'/>pt;
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
			</xsl:choose>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="w:pPr|w:tblPr|w:trPr|w:tcPr|w:sdtPr">
		<xsl:param name="scopeselector"/>
		<xsl:param name="listintend" />
		<xsl:variable name="nestedtdcssrules">
			<xsl:value-of select="$scopeselector"/> td {
			<xsl:for-each select="./*">
				<xsl:choose>
					<xsl:when test="local-name(.)='tblBorders'">
						<xsl:if test="count(../w:tblCellSpacing)=0 or number(../w:tblCellSpacing/@w) = 0">
							<xsl:call-template name="borders"/>
						</xsl:if>
					</xsl:when>
					<xsl:when test="local-name(.)='trHeight'">
						height:  <xsl:value-of select='number(@w:val) div 20'/>pt;
					</xsl:when>
				</xsl:choose>
			</xsl:for-each>
			 }
		</xsl:variable>
		<xsl:variable name="cssrules">
			<xsl:for-each select="./*">
				<xsl:choose>
					<xsl:when test="local-name(.)='tcW'">
						<xsl:choose>
							<xsl:when test="@w:type='pct' and contains(@w:w,'%')">
								width:<xsl:value-of select="@w:w"/>%;
							</xsl:when>
							<xsl:when test="@w:type='pct'">
								width:<xsl:value-of select="number(@w:w) div 50"/>%;
							</xsl:when>
							<xsl:when test="@w:type='dxa'">
								width:<xsl:value-of select="number(@w:w) div 12"/>pt;
							</xsl:when>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="local-name(.)='tcBorders'">
						<xsl:call-template name="borders" />
					</xsl:when>
					<xsl:when test="local-name(.)='spacing'">
						<xsl:for-each select="./@*">
							<xsl:choose>
								<xsl:when test='local-name(.)="after"'>
									margin-bottom: <xsl:value-of select='number(.) div 200'/>pt;
								</xsl:when>
								<xsl:when test='local-name(.)="before"'>
									margin-top: <xsl:value-of select='number(.) div 200'/>pt;
								</xsl:when>
							</xsl:choose>
						</xsl:for-each>
					</xsl:when>
					<xsl:when test="local-name(.)='pBdr'">
						<xsl:call-template name="borders"/>
					</xsl:when>
					<xsl:when test="local-name(.)='tblBorders'">
						<xsl:choose>
							<xsl:when test="count(../w:tblCellSpacing)=0 or number(../w:tblCellSpacing/@w) = 0">
								border-collapse:collapse;
							</xsl:when>
							<xsl:otherwise>
								<xsl:call-template name="borders"/>
							</xsl:otherwise>
						</xsl:choose>
						
					</xsl:when>
					<xsl:when test="local-name(.)='shd'">
						background-color:#<xsl:value-of select="@w:fill"/>;
					</xsl:when>
					<xsl:when test="local-name(.)='jc'">
						<xsl:choose>
							<xsl:when test="@w:val ='end'">
								text-align:right;
							</xsl:when>
							<xsl:when test="@w:val ='start'">
								text-align:left;
							</xsl:when>
							<xsl:when test="@w:val ='both'">
								text-align:justify;
							</xsl:when>
							<xsl:otherwise>
								text-align:<xsl:value-of select="@w:val"/>;
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="local-name(.)='ind' and local-name(../..)!='lvl'">
						<xsl:apply-templates select="current()">
							<xsl:with-param name="listintend" select="$listintend" />
						</xsl:apply-templates>
					</xsl:when>
				</xsl:choose>
			</xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="local-name(..)='style' or local-name(..)='lvl'"><xsl:value-of select="$cssrules"/></xsl:when>
			<xsl:when test="local-name(.)='sdtPr'"><xsl:value-of select="$scopeselector"/> { <xsl:apply-templates select="w:rPr" /> }</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name="style">
					<xsl:value-of select="$cssrules"/>
				</xsl:attribute>
				<style>
					<xsl:attribute name="scoped"/>
					<xsl:choose>
						<xsl:when test="local-name(.)='sdtPr'">
							<xsl:value-of select="$scopeselector"/> { <xsl:apply-templates select="w:rPr" /> }
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="$scopeselector"/> span { <xsl:apply-templates select="w:rPr" /> }
						</xsl:otherwise>
					</xsl:choose>
					<xsl:value-of select="$nestedtdcssrules"/>
				</style>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="w:body">
		<meta name="{'generator'}" content="{'docx2html.xsl https://github.com/ottoville/DOCX2HTML.XSLT'}"/>
		<article>
			<style>
				ins {text-decoration:none;}
				ul {list-style-position:inside}
				ul>li>p:first-child {display:inline;}
				ul.outsidee { list-style-position:outside} ul.outsidee>li>p:first-child { text-indent:0 !important;display:block}
				ul>li>p:not(:first-child) { text-indent:0 }
				ul, li { margin:0;padding:0 } li p {} input[type="text"] {height:18px} p {margin:0}
				article>* {display:none} article>section {display:block; margin-top:0.1cm;border-width:1px;border-style:solid;}
				<xsl:value-of select="base-uri(.)"/>
				<xsl:for-each select="document(resolve-uri('styles.xml',base-uri()))/w:styles/w:style">
					<xsl:choose>
						<xsl:when test="@w:type='paragraph'">
							p.<xsl:value-of select="./@w:styleId"/> { <xsl:apply-templates select="w:pPr" /> }
							p.<xsl:value-of select="./@w:styleId"/> span { <xsl:apply-templates select="w:rPr" /> }
						</xsl:when>
						<xsl:when test="@w:type='character'">
							span.<xsl:value-of select="./@w:styleId"/> { <xsl:apply-templates select="w:rPr" /> }
						</xsl:when>
					</xsl:choose>
				</xsl:for-each>
				<xsl:for-each select="document(resolve-uri('numbering.xml',base-uri()))/w:numbering/w:num">
					<xsl:variable name="listid" select="@w:numId"/>
					<xsl:variable name="absid" select="w:abstractNumId/@w:val"/>
					<xsl:for-each select="../w:abstractNum[@w:abstractNumId=$absid]/w:lvl[w:pPr|w:rPr]">
						<xsl:variable name="deep" select="number(@w:ilvl)+1"/>
						section ul.level<xsl:value-of select="@w:ilvl"/>.list<xsl:value-of select="$listid"/> { 
							<xsl:if test="w:pPr/w:ind/@w:left">
								padding-left:<xsl:value-of select='number(w:pPr/w:ind/@w:left) div 20'/>pt;
							</xsl:if>
						}
						section ul.level<xsl:value-of select="@w:ilvl"/>.list<xsl:value-of select="$listid"/>&gt;li {
							<xsl:if test="w:pPr/w:ind/@w:hanging">
								text-indent:-<xsl:value-of select='number(w:pPr/w:ind/@w:hanging) div 20'/>pt;
							</xsl:if>
						}
						section ul.level<xsl:value-of select="@w:ilvl"/>.list<xsl:value-of select="$listid"/>&gt;li&gt;p:first-child {
							<xsl:apply-templates select="w:pPr" /> 
						}
						section ul.level<xsl:value-of select="@w:ilvl"/>.list<xsl:value-of select="$listid"/>&gt;li&gt;p:first-child&gt;span { <xsl:apply-templates select="w:rPr" /> }
					</xsl:for-each>
				</xsl:for-each>
			</style>
			<!-- Match always last paragraph of page -->
			<xsl:for-each select="*[w:pPr/w:sectPr or (w:r/w:br/@w:type='page' and not(./following-sibling::*[1][w:pPr/w:sectPr]))]|w:sectPr">
				<xsl:call-template name="sections">
					<xsl:with-param name="reldocument" select="resolve-uri('_rels/document.xml.rels',base-uri())" />
				</xsl:call-template>
			</xsl:for-each>
		</article>
	</xsl:template>
	<xsl:template name="section">
		<xsl:param name="reldocument" />
		<xsl:param name="selector" />
		<xsl:for-each select="$selector">
			<xsl:variable name="listid" select="w:pPr/w:numPr/w:numId/@w:val" />
			<xsl:variable name="nextlistid" select="./following-sibling::w:p[w:pPr/w:numPr][not(w:pPr/w:numPr/w:numId/@w:val=$listid)][1]/w:pPr/w:numPr/w:numId/@w:val" />
			<xsl:if test="not(./preceding-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=$nextlistid]) and (not(w:pPr/w:numPr) or not(./preceding-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=current()/w:pPr/w:numPr/w:numId/@w:val]))">
				<xsl:apply-templates select="current()">
					<xsl:with-param name="reldocument" select="$reldocument" />
				</xsl:apply-templates>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="sections"><!-- Always the last page of document -->
		<xsl:param name="reldocument" />
		<!-- Either the page section, following section or document section-->
		<xsl:variable name="sectionselector" select="w:pPr/w:sectPr|./following-sibling::*[w:pPr/w:sectPr][1]/w:pPr/w:sectPr|/w:document/w:body/w:sectPr"/>
		<xsl:variable name="prevpage" select="preceding-sibling::*[w:pPr/w:sectPr or (w:r/w:br/@w:type='page' and not(./following-sibling::*[1][w:pPr/w:sectPr]))][1]" />
		<xsl:variable name="thisid" select="generate-id(current())" />
		<xsl:variable name="paddingtop" select="number($sectionselector[1]/w:pgMar/@w:top)" />
		<xsl:variable name="paddingbottom" select="number($sectionselector[1]/w:pgMar/@w:bottom)" />
		<xsl:variable name="paddingright" select="number($sectionselector[1]/w:pgMar/@w:right)" />
		<xsl:variable name="paddingleft" select="number($sectionselector[1]/w:pgMar/@w:left)" />
		<section>
			<xsl:attribute name="style">
				width:<xsl:value-of select="(number($sectionselector[1]/w:pgSz/@w:w) - $paddingright - $paddingleft) div 20"/>pt;
				min-height:<xsl:value-of select="(number($sectionselector[1]/w:pgSz/@w:h) - $paddingtop - $paddingbottom) div 20"/>pt;
				padding:<xsl:value-of select="$paddingtop div 20"/>pt <xsl:value-of select="$paddingright div 20"/>pt <xsl:value-of select="$paddingbottom div 20"/>pt <xsl:value-of select="$paddingleft div 20"/>pt;
			</xsl:attribute>
			<xsl:choose>
				<xsl:when test="not($prevpage)"> <!--first page or document with one page -->
					<xsl:if test='count($sectionselector[1]/w:headerReference[@w:type="first"])'>
						<header>
							<xsl:variable name="hyperlinkid" select='$sectionselector[1]/w:headerReference[@w:type="first"]/@r:id'/>
							<xsl:attribute name="id"><xsl:value-of select="$hyperlinkid" /></xsl:attribute>
							<xsl:apply-templates select="document(resolve-uri(document($reldocument)/*/*[@Id=$hyperlinkid]/@Target,base-uri()))/*/*">
								<xsl:with-param name="reldocument" select="resolve-uri(concat('_rels/',document($reldocument)/*/*[@Id=$hyperlinkid]/@Target,'.rels'),base-uri())" />
							</xsl:apply-templates>
						</header>
					</xsl:if>
					<xsl:call-template name="section">
						<xsl:with-param name="reldocument" select="$reldocument" />
						<xsl:with-param name="selector" select="preceding-sibling::*|current()" />
					</xsl:call-template>
				</xsl:when>
				<xsl:when test="w:pPr/w:sectPr or w:r/w:br/@w:type='page'">
					<xsl:call-template name="section">
						<xsl:with-param name="reldocument" select="$reldocument" />
						<xsl:with-param name="selector" select="current()|./preceding-sibling::*[count(./preceding-sibling::*) &gt; count($prevpage/preceding-sibling::*)]" />
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:call-template name="section">
						<xsl:with-param name="reldocument" select="$reldocument" />
						<xsl:with-param name="selector" select="$prevpage/following-sibling::*" />
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</section>
	</xsl:template>
	<xsl:template match="w:sdt">
		<xsl:param name="reldocument" />
		<xsl:param name="thisid" select="generate-id(current())" />
		<style>
			<xsl:apply-templates select="w:sdtPr">
				<xsl:with-param name="scopeselector" select="concat('.',$thisid)" />
			</xsl:apply-templates>
		</style>
		<xsl:apply-templates select="w:sdtContent">
			<xsl:with-param name="reldocument" select="$reldocument" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="w:sdtContent">
		<xsl:param name="reldocument" />
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="w:hyperlink">
		<xsl:param name="reldocument" />
		<a>
			<xsl:variable name="hyperlinkid" select="@r:id"/>
			<xsl:attribute name="href"><xsl:value-of select="document($reldocument)/*/*[@Id=$hyperlinkid]/@Target"/></xsl:attribute>
			<xsl:attribute name="target">_blank</xsl:attribute>
			<xsl:apply-templates />
		</a>
	</xsl:template>
	<xsl:template match="w:br">
		<br>
			<xsl:if test="@w:type='page'">
				<xsl:attribute name="style">page-break-before:always</xsl:attribute>
			</xsl:if>	
		</br>
	</xsl:template>
	<xsl:template match="w:instrText">
	</xsl:template>
	<xsl:template match="w:ins">
		<xsl:param name="reldocument" />
		<xsl:param name="listintend" />
		<ins>
			<xsl:apply-templates select="w:pPr|w:r">
				<xsl:with-param name="listintend" select="$listintend" />
				<xsl:with-param name="reldocument" select="$reldocument" />
			</xsl:apply-templates>
		</ins>
	</xsl:template>
	
	<xsl:template match="pic:pic">
		<xsl:param name="reldocument" />
		<xsl:param name="position" />
		<xsl:variable name="dravingid" select="pic:blipFill/a:blip/@r:embed"/>
		<img>
			<xsl:attribute name="style">display:inline;width:<xsl:value-of select="number($position/@cx) div 360000"/>cm;height:<xsl:value-of select="number($position/@cy) div 360000"/>cm</xsl:attribute>
			<xsl:attribute name="src"><xsl:value-of select="resolve-uri(document($reldocument)/*/*[@Id=$dravingid]/@Target,base-uri())"/></xsl:attribute>
		</img>
	</xsl:template>
	<xsl:template match="wps:wsp">
		<svg>
		</svg>
	</xsl:template>
	<xsl:template match="wp:anchor">
		<xsl:param name="reldocument" />
		<xsl:variable name="dravingid" select="a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed"/>
		<xsl:attribute name="style">width:<xsl:value-of select="number(wp:extent/@cx) div 360000"/>cm;height:<xsl:value-of select="number(wp:extent/@cy) div 360000"/>cm</xsl:attribute>
		<xsl:attribute name="src"><xsl:value-of select="document($reldocument)/*/*[@Id=$dravingid]/@Target"/></xsl:attribute>
	</xsl:template>
	<xsl:template match="wp:inline">
		<xsl:param name="reldocument" />		
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="position" select="wp:extent" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="a:graphic">
		<xsl:param name="reldocument" />		
		<xsl:param name="position" />		
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="position" select="$position" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="a:graphicData">
		<xsl:param name="reldocument" />		
		<xsl:param name="position" />		
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
			<xsl:with-param name="position" select="$position" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="w:drawing">
		<xsl:param name="reldocument" />
		<xsl:apply-templates>
			<xsl:with-param name="reldocument" select="$reldocument" />
		</xsl:apply-templates>
	</xsl:template>
</xsl:stylesheet>
