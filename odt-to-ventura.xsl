<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"   
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" 
    exclude-result-prefixes="office text style fo"
    >

    <xsl:output method="text" encoding="WINDOWS-1251"/>

    <xsl:key name="auto-styles" match="style:style" use="@style:name"/>
    <xsl:key name="paragraph-styles" match="text:h | text:p" use="@text:style-name"/>
    <xsl:key name="paragraph-parent-styles" match="style:style"  
        use="@style:parent-style-name"/>

    <xsl:variable name="styles_file" select="'styles.xml'"/>
    <xsl:variable name="content_file" select="/"/>

    <xsl:template match="/">
        <xsl:apply-templates select="/office:document-content"/>
    </xsl:template>
    
    <!-- =================== main() ===================================== -->
    <xsl:template match="/office:document-content">
        <xsl:text>@Z_STYLE70 = </xsl:text>
        <!-- main processing -->
        <xsl:apply-templates select="office:body/office:text"/>
    </xsl:template>

    <xsl:template match="office:text">
        <xsl:apply-templates select="table:table | text:h | text:list | text:p"/>
    </xsl:template>


    <!-- ===========================list processing =============  -->
    <xsl:template match="text:list">
        <xsl:apply-templates select="descendant::text:list-item"/>
    </xsl:template>

    <xsl:template match="text:list-item">
        <!-- here we could do something to preserve bullets and numbers but
             instead we just do nothing here -->
        <xsl:apply-templates select="text:p | text:h"/>
    </xsl:template>


    <!-- ========================== text processing =================== -->
    <xsl:template match="text:p | text:h">
        <!-- processing main para tags -->
        <xsl:text>&#xD;&#xA;@</xsl:text>

        <xsl:call-template name="getParaStyle">
            <xsl:with-param name="style-name" select="@text:style-name"/>
        </xsl:call-template>

        <xsl:text> = </xsl:text>
        <xsl:apply-templates/>
    </xsl:template>

    <xsl:template match="text:span">
        <!-- processing span tags -->
        
        <xsl:call-template name="getCustomFormatting"/>
        <xsl:apply-templates/>
        <xsl:call-template name="getCustomFormatting">
            <xsl:with-param name="clear" select="'1'"/>
        </xsl:call-template>
    </xsl:template>

    <!-- Index marks and fields -->
    <xsl:template match="text:hidden-text">
        <!-- do nothing, dont take index marks -->
    </xsl:template>
    
    <xsl:template match="text:alphabetical-index-mark">
        <!-- index mark -->
        <!-- what we receive:
            <text:alphabetical-index-mark text:string-value="содержимое индекса{123}" text:key1="Рубрика" text:key2="Подраздел"/>
            or
            <text:alphabetical-index-mark text:string-value="Cодержимое индекса{123}" text:key1="" text:key2=""/>
             what we create here:
             <$IРубрика:подрубрика:содержание{123}>
        -->

        <xsl:text>&lt;$I</xsl:text>
        <!-- lets break apart the 3 levels -->
        <xsl:if test="@text:key1">
            <xsl:value-of select="@text:key1"/>
            <xsl:text>:</xsl:text>
        </xsl:if>
        <xsl:if test="@text:key2">
            <xsl:value-of select="@text:key2"/>
            <xsl:text>:</xsl:text>
        </xsl:if>
        <xsl:value-of select="@text:string-value"/>
        <xsl:text>&gt;</xsl:text>
    </xsl:template>

    <!-- Bookmark and bookmark reference aka see p. 11 -->
    <xsl:template match="text:bookmark-start|text:bookmark">
        <xsl:text>&lt;$M[</xsl:text>
        <xsl:value-of select="@text:name" />
        <xsl:text>]&gt;</xsl:text>
    </xsl:template>

    <xsl:template match="text:bookmark-ref">
        <xsl:text>&lt;$R[P#,</xsl:text>
        <xsl:value-of select="@text:ref-name" />
        <xsl:text>,1,,0,]&gt;</xsl:text>
    </xsl:template>

    <xsl:template match="text:line-break">
        <xsl:text>&lt;R&gt;</xsl:text>
    </xsl:template>
    <xsl:template match="text:tab">
        <xsl:text>&#09;</xsl:text>
    </xsl:template>
    <xsl:template match="text:s">
            <xsl:call-template name="printWhitespaces">
                <xsl:with-param name="count" select="number(@text:c)"/>
            </xsl:call-template>
    </xsl:template>


    <!-- ======================= FINISHED text processing ============== -->

    <!-- ======================= table processing ===================== -->
    
    
    <xsl:template match="table:table">
        <!-- processing table start tags 
        -->
        <!-- we need to count columns right -->
        <xsl:variable name="colCount" select="count(table:table-column[not(@table:number-columns-repeated)])"/>
        <xsl:variable name="repeatedColSum">
            <xsl:value-of select="sum(table:table-column/@table:number-columns-repeated)"/>
        </xsl:variable>

        <xsl:text>&#xD;&#xA;@Z_TBL_BEG = VERSION(10), TAGNAME(Default Table), ROWS(</xsl:text>
        <xsl:value-of select="count(table:table-row)"/>
        <xsl:text>), COLUMNS(</xsl:text>
        <xsl:value-of select="$colCount + $repeatedColSum"/>
        <xsl:text>)</xsl:text>

        <!-- go process the rows -->
        <xsl:apply-templates select="table:table-row"/>
        <!-- finishing table -->
        <xsl:text>&#xD;&#xA;@Z_TBL_END = </xsl:text>
    </xsl:template>

    <xsl:template match="table:table-row">
        <xsl:text>&#xD;&#xA;@Z_TBL_ROW_BEG = </xsl:text>
        <xsl:apply-templates select="table:table-cell | table:covered-table-cell">
        </xsl:apply-templates>
        <xsl:text>&#xD;&#xA;@Z_TBL_ROW_END = </xsl:text>
    </xsl:template>

    <xsl:template match="table:table-cell">
            <!-- we are in table-cell -->
            <xsl:text>&#xD;&#xA;@Z_TBL_CELL_BEG = </xsl:text>
            <xsl:apply-templates/>
            <xsl:text>&#xD;&#xA;@Z_TBL_CELL_END = </xsl:text>
    </xsl:template>
    <xsl:template match="table:covered-table-cell">
        <!-- Got merged cell. Need to find out if it is HJOINED or VJOINED -->
        <xsl:variable name="rownum" select="count(parent::table:table-row/preceding-sibling::*)+1" />
        <xsl:variable name="preceding-rows" select="parent::table:table-row/preceding-sibling::table:table-row"/>
        <xsl:variable name="colnum" select="count(preceding-sibling::*)" /> 
        <!-- if there is some above cell in the same column that has rowspan value greater or equal amount of rows between sayed cells row and current rownum than it is vjoined, otherwise it is hjoined cell -->
        <xsl:text>&#xD;&#xA;</xsl:text>
        <xsl:text>@Z_TBL_CELL_BEG = </xsl:text>
        <xsl:choose>
            <xsl:when test="$preceding-rows/table:table-cell[count(preceding-sibling::*) = $colnum and number(@table:number-rows-spanned) &gt; count(parent::*/following-sibling::*[count(.|$preceding-rows) = count($preceding-rows)])+1]">
                <xsl:text>VJOINED</xsl:text> 
            </xsl:when>
            <xsl:otherwise>
                <xsl:text>HJOINED</xsl:text>
            </xsl:otherwise>
        </xsl:choose>
        <xsl:text>&#xD;&#xA;</xsl:text>
        <xsl:text>@Z_TBL_CELL_END = </xsl:text>
    </xsl:template>
    
    

    <!-- ======================= FINISHED table processing ===================== -->



    <!-- =========================== utilities =========================== -->
    <xsl:template name="printWhitespaces">
        <xsl:param name="count"/>
        <xsl:text> </xsl:text>
        <xsl:if test="$count &gt; 1">
            <xsl:call-template name="printWhitespaces">
                <xsl:with-param name="count" select="$count - 1"/>
            </xsl:call-template>
        </xsl:if>
    </xsl:template>

    <xsl:template name="getCustomFormatting">
        <!-- checking for accidental formatting like Bold, Subscript -->
        <xsl:param name="clear"/>

        <!-- first we find out the style name and the parent (Paragraph) style -->
        <xsl:variable name="styleName">
            <xsl:choose>
                <xsl:when test="self::text()">
                    <xsl:value-of select="parent::*[1]/@text:style-name"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="@text:style-name"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="parentStyleName">
            <xsl:value-of select="parent::*[1][self::text:p|self::text:h]/@text:style-name"/>
        </xsl:variable>

        <!-->
            <xsl:message>
                <xsl:text>tag is: </xsl:text>
                <xsl:value-of select="name(.)"/>
                <xsl:text>&#10;styleName: </xsl:text>
                <xsl:value-of select="$styleName"/>
                <xsl:text>&#10;parentStyleName: </xsl:text>
                <xsl:value-of select="$parentStyleName"/>
            </xsl:message>
        </!-->

        <!-- now we aquire values and pick the right (recent) ones -->
        <xsl:if test="key('auto-styles', $styleName)">
            <xsl:variable name="fname">
                <!-- checking for symbol, not implemented yet -->
                <xsl:value-of
                    select="key('auto-styles', $styleName)/style:text-properties/@style:font-name"/>
            </xsl:variable>  
            <xsl:variable name="fposition">
                <!-- subscript or super -->
                <xsl:value-of 
                    select="concat(
                    key('auto-styles', $styleName)/style:text-properties/@style:text-position,
                    key('auto-styles', $parentStyleName)/style:text-properties/@style:text-position)"/>
            </xsl:variable>
            <xsl:variable name="funderline">
                <xsl:value-of 
                    select="concat(
                    key('auto-styles', $styleName)/style:text-properties/@style:text-underline-style,
                    key('auto-styles', $parentStyleName)/style:text-properties/@style:text-underline-style)"/>
            </xsl:variable>
            <xsl:variable name="fweight">
                <xsl:value-of 
                    select="concat(
                    key('auto-styles', $styleName)/style:text-properties/@fo:font-weight,
                    key('auto-styles', $parentStyleName)/style:text-properties/@fo:font-weight)"/>
            </xsl:variable>
            <xsl:variable name="fstyle">
                <xsl:value-of 
                    select="concat(
                    key('auto-styles', $styleName)/style:text-properties/@fo:font-style,
                    key('auto-styles', $parentStyleName)/style:text-properties/@fo:font-style)"/>
            </xsl:variable>
            <xsl:variable name="fbold">
                <xsl:value-of select="starts-with($fweight, 'bold')"/>
            </xsl:variable>
            <xsl:variable name="fitalic">
                <xsl:value-of select="starts-with($fstyle,'italic') 
                    or starts-with($fstyle, 'oblique')"/>
            </xsl:variable>

            <xsl:variable name="font-style">
                <xsl:text>&lt;</xsl:text>
                <xsl:if test="$fbold='true' and not($clear)">
                    <xsl:text>B</xsl:text>
                </xsl:if>
                <xsl:if test="$fbold='true' and $clear">
                    <xsl:text>W0</xsl:text>
                </xsl:if>
                <xsl:if test="$fitalic='true'">
                    <xsl:text>I</xsl:text>
                    <xsl:if test="$clear"><xsl:text>*</xsl:text></xsl:if>
                </xsl:if>
                <xsl:if test="string-length($funderline)">
                    <xsl:if test="not(starts-with($funderline, 'none'))">
                        <xsl:text>U</xsl:text>
                        <xsl:if test="$clear"><xsl:text>*</xsl:text></xsl:if>
                    </xsl:if>
                </xsl:if>
                <xsl:if test="string-length($fposition)">
                    <xsl:if test="not($clear)">
                        <xsl:choose>
                            <xsl:when test="starts-with($fposition, 'sub') or 
                                number(substring-before($fposition,'%')) &lt; 0">
                                <xsl:text>V</xsl:text>
                            </xsl:when>
                            <xsl:when test="starts-with($fposition, 'super') or
                                number(substring-before($fposition,'%')) &gt; 0">
                                <xsl:text>^</xsl:text>
                            </xsl:when>
                        </xsl:choose>
                    </xsl:if>
                    <xsl:if test="$clear"><xsl:text>^*</xsl:text></xsl:if>
                </xsl:if>
                <xsl:text>&gt;</xsl:text>
            </xsl:variable>

            <!-- Here we print the aquired values -->

            <xsl:if test="string-length($font-style) &gt; 2">
                <xsl:value-of select="$font-style"/>
            </xsl:if>
            <xsl:if test="contains($fname, 'Symbol') and not($clear)">
                <xsl:text>&lt;F"Symbol"&gt;</xsl:text>
            </xsl:if>
            <xsl:if test="contains($fname, 'Symbol') and $clear">
                <xsl:text>&lt;F255&gt;</xsl:text>
            </xsl:if>

        </xsl:if>
    </xsl:template>

    <xsl:template name="getParaStyle">
        <!-- returns the style name or parent style name if auto-style -->
        <xsl:param name="style-name"/>
        <xsl:choose>
            
            <xsl:when test="key('auto-styles', $style-name)">
                <xsl:value-of 
                    select="key('auto-styles', $style-name)/@style:parent-style-name"/>
            </xsl:when>

            <xsl:otherwise>
                <xsl:value-of select="$style-name"/>
            </xsl:otherwise>

        </xsl:choose>
    </xsl:template>

    <xsl:template name="search-and-replace">
        <xsl:param name="input"/>
        <xsl:param name="search-string"/>
        <xsl:param name="replace-string"/>
        <xsl:choose>
            <!-- See if the input contains the search string -->
            <xsl:when test="$search-string and 
                            contains($input,$search-string)">
            <!-- If so, then concatenate the substring before the search
            string to the replacement string and to the result of
            recursively applying this template to the remaining substring.
            -->
                <xsl:value-of 
                        select="substring-before($input,$search-string)"/>
                <xsl:value-of select="$replace-string"/>
                <xsl:call-template name="search-and-replace">
                        <xsl:with-param name="input"
                        select="substring-after($input,$search-string)"/>
                        <xsl:with-param name="search-string" 
                        select="$search-string"/>
                        <xsl:with-param name="replace-string" 
                            select="$replace-string"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <!-- There are no more occurrences of the search string so 
                just return the current input string -->
                <xsl:value-of select="$input"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    <!-- ========================= FINISHED utilities =================== -->
        


    <!-- ============================ getting styles ====================   -->
    <xsl:template match="office:document-styles/office:styles/style:style">
        <!-- we should be inside the styles.xml now -->
        <xsl:variable name="style_name" >
            <xsl:value-of select="string(@style:name)"/>
        </xsl:variable>
        <xsl:apply-templates select="$content_file/*/office:body">
            <xsl:with-param name="style_name" select="$style_name"/>
        </xsl:apply-templates>
    </xsl:template>

    <xsl:template match="office:body">
        <!-- just looking for used styles here -->
        <xsl:param name="style_name"/>
        <xsl:choose>
            <xsl:when test="key('paragraph-styles', $style_name)">
                <xsl:text>&#xD;&#xA;&lt;DefineParaStyle:</xsl:text>
                <xsl:value-of select="$style_name"/>
                <xsl:text>=&lt;BasedOn:NormalParagraphStyle&gt;&gt;</xsl:text>
            </xsl:when>
            <xsl:otherwise>
                <xsl:if test="key('paragraph-parent-styles', $style_name)">
                    <xsl:text>&#xD;&#xA;&lt;DefineParaStyle:</xsl:text>
                    <xsl:value-of select="$style_name"/>
                    <xsl:text>&gt;</xsl:text>
                </xsl:if>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    <!--  ======================== FINISHED getting styles ================ -->
</xsl:stylesheet>
