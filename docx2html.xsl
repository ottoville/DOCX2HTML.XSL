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
    xmlns:pt="http://powertools.codeplex.com/2011"
    xmlns:d2h="http://www.github.com/rnathanday/docx2html"
    xmlns:ixsl="http://saxonica.com/ns/interactiveXSLT"
	xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
	xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    extension-element-prefixes="ixsl"
    exclude-result-prefixes="xs pt w r pr d2h ixsl wp a pic xhtml"
    version="2.0">
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
	<xsl:template match="w:rPr" name="spanstyle">
		<xsl:variable name="cssrules">
			<xsl:if test="../w:t/@xml:space='preserve'">white-space:pre-wrap;</xsl:if>
			<xsl:for-each select="./*">
				<xsl:choose>
					<xsl:when test="local-name(.)='spacing'">letter-spacing:<xsl:value-of select='number(@w:val) div 12'/>pt;</xsl:when>
					<xsl:when test='local-name(.)="kern" and number(../w:sz/@w:val) &gt;= number(@w:val)'>
						font-kerning:auto;
					</xsl:when>
					<xsl:when test="local-name(.)='shd'">background-color:#<xsl:value-of select="@w:fill"/>;</xsl:when>
					<xsl:when test="local-name(.)='rFonts'">font-family:<xsl:value-of select="@w:ascii|@w:cs|@w:eastAsia"/>;</xsl:when>
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
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="local-name(..)='style' or local-name(..)='pPr' or local-name(..)='lvl' or local-name(..)='sdtPr'">
				<xsl:value-of select="$cssrules"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name="style">
					<xsl:value-of select="$cssrules"/>
				</xsl:attribute>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="w:body">
		<article>
			<style>
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
			<xsl:for-each select="*[w:pPr/w:sectPr or ./following-sibling::*[1][w:r/w:br/@w:type='page'][not(./following-sibling::*[1][w:pPr/w:sectPr])]]|w:sectPr">
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
	<xsl:template name="sections"><!-- Aina sivun vimppa kappale -->
		<xsl:param name="reldocument" />
		<!-- Oma sivumääritelmä, tai seuraava, tai dokumentin-->
		<xsl:variable name="sectionselector" select="w:pPr/w:sectPr|./following-sibling::*[w:pPr/w:sectPr][1]/w:pPr/w:sectPr|/w:document/w:body/w:sectPr"/>
		<xsl:variable name="prevpage" select="preceding-sibling::*[w:pPr/w:sectPr or ./following-sibling::*[1][w:r/w:br/@w:type='page'][not(./following-sibling::*[1][w:pPr/w:sectPr])]][1]" />
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
				<xsl:when test="not($prevpage)"> <!--eka sivu tai yksisivuinen asiakirja -->
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
				<xsl:when test="w:pPr/w:sectPr or ./following-sibling::*[1][w:r/w:br/@w:type='page']">
					<xsl:call-template name="section">
						<xsl:with-param name="reldocument" select="$reldocument" />
						<xsl:with-param name="selector" select="./preceding-sibling::*[count(./preceding-sibling::*) &gt; count($prevpage/preceding-sibling::*)]" />
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
				<xsl:apply-templates select="w:tcPr|w:p">
					<xsl:with-param name="reldocument" select="$reldocument" />
				</xsl:apply-templates>
			</td>
		</xsl:if>
	</xsl:template>
	<xsl:template match="w:tr">
		<xsl:param name="reldocument" />
		<tr>
			<xsl:apply-templates select="w:trPr|w:tc">
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
	<!-- Lista jatkuu-->
	<xsl:template name="listitem" match="w:p[w:pPr/w:numPr and count(./preceding-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=current()/w:pPr/w:numPr/w:numId/@w:val]) &gt; 0 ]">
		<xsl:param name="reldocument" />
		<xsl:param name="listintend"/>
		<xsl:param name="thismargin">
			<xsl:choose>
				<xsl:when test="w:pPr/w:ind/@w:left"><xsl:value-of select="number(w:pPr/w:ind/@w:left)"/></xsl:when>
				<xsl:otherwise><xsl:value-of select="number(0)"/></xsl:otherwise>
			</xsl:choose>
		</xsl:param>
		<xsl:variable name="listid" select="w:pPr/w:numPr/w:numId/@w:val" />
		<xsl:variable name="currentid" select="generate-id(current())" />
		<xsl:variable name="nextlistid" select="following-sibling::w:p[w:pPr/w:numPr][not(w:pPr/w:numPr/w:numId/@w:val=$listid)][1]/w:pPr/w:numPr/w:numId/@w:val" />
		<xsl:variable name="prevlistid" select="preceding-sibling::w:p[w:pPr/w:numPr][not(w:pPr/w:numPr/w:numId/@w:val=$listid)][1]/w:pPr/w:numPr/w:numId/@w:val" />
		<li>
			<xsl:attribute name="style">margin-left:<xsl:value-of select='number($thismargin) div 20'/>pt;
				<xsl:if test="w:pPr/w:ind/@w:hanging">
					text-indent:-<xsl:value-of select='number(w:pPr/w:ind/@w:hanging) div 20'/>pt;
				</xsl:if>
			</xsl:attribute>
			<xsl:attribute name="jatkuvalista"><xsl:value-of select="$listintend"/></xsl:attribute>
			<xsl:attribute name="id"><xsl:value-of select="generate-id(current())"/></xsl:attribute>
			<xsl:attribute name="prevcount"><xsl:value-of select="count(./preceding-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=current()/w:pPr/w:numPr/w:numId/@w:val])"/></xsl:attribute>
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
	<!-- Listan ulkopuoliset kappaleet-->
	<xsl:template name="paragraph" match="w:p[not(w:pPr/w:numPr)]">
		<xsl:param name="reldocument" />
		<xsl:param name="listintend"/>
		<xsl:variable name="class">
			<xsl:if test="count(w:pPr/w:pStyle)">
				<xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
			</xsl:if>
			<xsl:if test="local-name(..)='sdtContent'">
				<xsl:value-of select="concat(' ',generate-id(../..))"/>
			</xsl:if>
		</xsl:variable>
		<p class="{normalize-space($class)}" pid="{generate-id(.)}">
			<xsl:apply-templates select="w:pPr|w:r">
				<xsl:with-param name="scopeselector">p[pid='<xsl:value-of select="generate-id(.)"/>']</xsl:with-param>
				<xsl:with-param name="listintend" select="$listintend" />
				<xsl:with-param name="reldocument" select="$reldocument" />
			</xsl:apply-templates>
			<xsl:if test='count(w:r)=0'>
				<br />
			</xsl:if>
		</p>
	</xsl:template>
	<!-- Lista alkaa-->
	<xsl:template match="w:p[w:pPr/w:numPr and not(./preceding-sibling::w:p[w:pPr/w:numPr/w:numId/@w:val=current()/w:pPr/w:numPr/w:numId/@w:val])]">
		<xsl:param name="reldocument" />
		<xsl:param name="listid" select="w:pPr/w:numPr/w:numId/@w:val" />
		<xsl:param name="listlevel" select="w:pPr/w:numPr/w:ilvl/@w:val" />
		<xsl:param name="relid" select="document(resolve-uri('numbering.xml',base-uri()))/w:numbering/w:num[@w:numId=$listid]/w:abstractNumId/@w:val" />
		<xsl:param name="thismargin">
			<xsl:choose>
				<xsl:when test="w:pPr/w:ind/@w:left"><xsl:value-of select="number(w:pPr/w:ind/@w:left)"/></xsl:when>
				<xsl:otherwise><xsl:value-of select="number(0)"/></xsl:otherwise>
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
				<xsl:with-param name="listintend" select="number($thislistintend + $thismargin)" />
			</xsl:call-template>
			<xsl:apply-templates select="./following-sibling::w:p[(
																		w:pPr/w:numPr and
																		w:pPr/w:numPr/w:numId/@w:val=$listid)]">
				<xsl:with-param name="reldocument" select="$reldocument" />
				<xsl:with-param name="listintend" select="number($thislistintend)" />
			</xsl:apply-templates>
		</ul>
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
	<xsl:template match="w:tab">
		<xsl:choose>
			<xsl:when test="count(../../w:pPr/w:ind[@w:hanging]) and count(./preceding-sibling::w:r/w:tab)=0">
				<br style="{'width:18pt;display:inline-block'}"/>
			</xsl:when>
			<xsl:otherwise>
				<span><xsl:attribute name="style">width:18pt;display:inline-block</xsl:attribute> </span>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="w:t">
		<xsl:value-of select="."/>
	</xsl:template>
	<xsl:template match="w:instrText">
	</xsl:template>
	<xsl:template match="w:r">
		<xsl:param name="reldocument" />
		<xsl:variable name="lastend" select="count(./preceding-sibling::w:r/w:fldChar[@w:fldCharType='end']/../preceding-sibling::w:r)" />
		<xsl:variable name="laststart" select="count(./preceding-sibling::w:r/w:fldChar[@w:fldCharType='separate']/../preceding-sibling::w:r)" />
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
		<xsl:if test="$laststart = 0 or $lastend &gt; $laststart">
			<xsl:choose>
				<xsl:when test="count(w:br) and count(*)=1">
					<xsl:apply-templates/>
				</xsl:when>
				<xsl:otherwise>
					<span>
						<xsl:attribute name="class" select="normalize-space($class)"/>
						<xsl:apply-templates select="w:rPr|w:t|w:fldChar|w:drawing">
							<xsl:with-param name="reldocument" select="$reldocument" />
						</xsl:apply-templates>
					</span>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
	</xsl:template>
	<xsl:template match="w:fldChar[@w:fldCharType='begin']">
		<xsl:variable name="fieldtype" select="normalize-space(../following-sibling::w:r[1]/w:instrText)" />
		<xsl:variable name="textstart" select="count(../following-sibling::w:r/w:fldChar[@w:fldCharType='separate']/../preceding-sibling::w:r)" />
		<xsl:variable name="textend" select="count(../following-sibling::w:r/w:fldChar[@w:fldCharType='end']/../preceding-sibling::w:r)" />
		<xsl:variable name="type">
			<xsl:choose>
				<xsl:when test="$fieldtype='FORMCHECKBOX'">
					<xsl:value-of select="'checkbox'" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="'text'" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="placeholder">
			<xsl:apply-templates select="../following-sibling::w:r[position() &gt; $textstart][position() &lt; $textend]/w:t" />
		</xsl:variable>
		<input type="{$type}" placeholder="{$placeholder}"/>
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
