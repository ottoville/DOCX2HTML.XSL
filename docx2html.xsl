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
	<!-- docx2html.xsl
		Project started by Otto-Ville Lamminpää
			ottoville.lamminpaa@gmail.com
			+358445596869
	-->
	<xsl:import href="pages.xsl"/>
	<xsl:import href="paragraphs.xsl"/>
	<xsl:import href="lists.xsl"/>
	<xsl:import href="text.xsl"/>
	<xsl:import href="table.xsl"/>
	<xsl:import href="images.xsl"/>
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
					<!--<xsl:attribute name="scoped"/>-->
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
	<xsl:template match="w:style[@w:type='paragraph']">
		<xsl:if test="@w:default=1">
			<xsl:value-of select="concat('p.',' { ')"/><xsl:apply-templates select="w:pPr" /> }
			<xsl:value-of select="concat('p>span { ')"/><xsl:apply-templates select="w:rPr" /> }
			<xsl:for-each select="w:pPr/w:tabs/w:tab">
				<xsl:value-of select="concat('p div.tab:nth-of-type(',position(),') { width: ',(number(./@w:pos) div 20) * (4 div 3),'px !important }')"/>
			</xsl:for-each>	
		</xsl:if>
		<xsl:value-of select="concat('p.',./@w:styleId,' { ')"/><xsl:apply-templates select="w:pPr" /> }
		<xsl:value-of select="concat('p.',./@w:styleId,'>span { ')"/><xsl:apply-templates select="w:rPr" /> }
		<xsl:for-each select="w:pPr/w:tabs/w:tab">
			<xsl:value-of select="concat('p.',../../../@w:styleId,' div.tab:nth-of-type(',position(),') { width: ',(number(./@w:pos) div 20) * (4 div 3),'px !important }')"/>
		</xsl:for-each>				
	</xsl:template>
	<xsl:template match="w:style[@w:type='character']">
		<xsl:if test="@w:default=1">
			<xsl:value-of select="concat('span { ')"/><xsl:apply-templates select="w:rPr" /> }
		</xsl:if>
		<xsl:value-of select="concat('span.',./@w:styleId,' { ')"/><xsl:apply-templates select="w:rPr" /> }
	</xsl:template>
	<xsl:template match="w:body">
		<xsl:variable name="themefile" select="document(resolve-uri(document(resolve-uri('_rels/document.xml.rels',base-uri()))/*/*[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme']/@Target,base-uri()))"/>

		<meta name="{'generator'}" content="{'docx2html.xsl https://github.com/ottoville/DOCX2HTML.XSLT'}"/>
		<article>
			<!--	Match always last paragraph of page.
					Cannot use apply-templates here because it would cause endles loop as 
					same paragraphs are going to be iterated in section template
			-->
			<xsl:for-each select="*[w:pPr/w:sectPr]|w:sectPr">
				<xsl:call-template name="sections">
					<xsl:with-param name="themefile" select="$themefile" />
				</xsl:call-template>
			</xsl:for-each>
		</article>
			<style>
				.tab {display:inline-block;text-indent:0;}
				ins {text-decoration:none;}
				ul {list-style-position:inside}
				ul>li>p:first-child {display:inline;}
				ul.outsidee { list-style-position:outside} ul.outsidee>li>p:first-child { text-indent:0 !important;display:block}
				ul>li>p:not(:first-child) { text-indent:0 }
				ul, li { margin:0;padding:0 } li p {} input[type="text"] {height:18px} input[type="checkbox"] { margin:0 } p {margin:0;position:relative}
				article {
					counter-reset: page;
				}
				article>div {
					counter-increment: page;
					display:block; margin-top:0.1cm;border-width:1px;border-style:solid;position:relative;}
				article>div output.pagenumber::before {
					content: counter(page);
				}
				article>div>header.even {
					display:none;
				}
				article>div:nth-of-type(even)>header.even {
					display:block;
				}
				article>div:nth-of-type(even)>header.even {
					display:block;
				}
				article>div:nth-of-type(even)>header.even+header.default {
					display:none;
				}
				article>div>header.first {
					display:none;
				}
				article>hr[data-firstheader="true"]+div>header.first {
					display:block;
				}
				article>hr[data-firstheader="true"]+div>header.first ~ header {
					display:none !important;
				}
				
				article>div>footer {
					position:absolute;
					bottom:0;
				}
				article>div>footer.even {
					display:none;
				}
				article>div:nth-of-type(even)>footer.even {
					display:block;
				}
				article>div:nth-of-type(even)>footer.even {
					display:block;
				}
				article>div:nth-of-type(even)>footer.even+footer.default {
					display:none;
				}
				article>div>footer.first {
					display:none;
				}
				article>hr[data-firstheader="true"]+div>footer.first {
					display:block;
				}
				article>hr[data-firstheader="true"]+div>footer.first ~ footer {
					display:none !important;
				}
				<xsl:apply-templates select="document(resolve-uri('styles.xml',base-uri()))/w:styles/w:style[not(@w:type='table')]" />
				<xsl:apply-templates select="document(resolve-uri('numbering.xml',base-uri()))/w:numbering/w:num" />
				<xsl:for-each select="$themefile/a:theme/a:themeElements/a:clrScheme/*">
					.<xsl:value-of select="local-name(.)"/>fill {
						<xsl:for-each select="./*">
							<xsl:choose>
								<xsl:when test="local-name(.)='srgbClr'">
									fill:#<xsl:value-of select="@val"/>;
								</xsl:when>
							</xsl:choose>
						</xsl:for-each>
					}
					.<xsl:value-of select="local-name(.)"/>border {
						<xsl:for-each select="./*">
							<xsl:choose>
								<xsl:when test="local-name(.)='srgbClr'">
									stroke:#<xsl:value-of select="@val"/>;
								</xsl:when>
							</xsl:choose>
						</xsl:for-each>
					}
				</xsl:for-each>
			</style>
		<script>
			<xsl:text>
				<![CDATA[
				var tabs=document.getElementsByClassName('tab');
				for(t=0;t<tabs.length;t++) {
					var tabinterval=parseFloat(tabs[t].style.minWidth);
					var realwidth=parseFloat(window.getComputedStyle(tabs[t]).width)
					if(realwidth>tabinterval) {
						var tabnumber=Math.ceil(realwidth/tabinterval);
						for(tt=tabnumber-1,next=tabs[t].nextSibling;	tt>0&&next&&next.classList.contains('tab')&&!next.firstChild	;next=next.nextSibling,tt--) {
							next.parentNode.removeChild(next);
						}
						tabs[t].style.width=tabnumber*tabinterval+'px';	
					}
				}
				]]>
			</xsl:text>
		</script>
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
</xsl:stylesheet>
