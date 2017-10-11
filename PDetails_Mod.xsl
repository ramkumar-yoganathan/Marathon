<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   <xsl:template match="/">
      <HTML>
         <HEAD>
            <META HTTP-EQUIV="Content-Type" CONTENT="text/html" />
         </HEAD>

         <BODY TEXT="#000000" BGCOLOR="#FFFFFF" LINK="#0000CC" VLINK="#660066" ALINK="#660066">
            <CENTER>
               <H1>
                  <FONT COLOR="#000050">
                     <FONT FACE="Times New Roman">
                  		<xsl:value-of select="//Report/Doc/DName" />
   					 </FONT>
                  </FONT>
               </H1>

               <!--TABLE BORDER="1" BGCOLOR="#FFFFFF" CELLSPACING="0" BORDERCOLOR="#808080" CELLPADDING="6">
                  <TR>
                     <TD class="hl0">
                        <FONT COLOR="Red">Failed : 
                        <xsl:value-of select="//Report/Doc/Summary/@failed" />
                        </FONT>
                     </TD>

                     <TD class="hl0">
                        <FONT COLOR="Orange">Warnings : 
                        <xsl:value-of select="//Report/Doc/Summary/@warnings" />
                        </FONT>
                     </TD>
                  </TR>
               </TABLE-->

               <h3>
                  <FONT FACE="Times New Roman">Execution started : 
                  <xsl:value-of select="//Report/Doc/Summary/@sTime" />
                  </FONT>
               </h3>
            </CENTER>

            <CENTER>
               <TABLE BORDER="1" BGCOLOR="#FFFFFF" CELLSPACING="0" BORDERCOLOR="#808080" CELLPADDING="6">
                  <TR ALIGN="CENTER" BGCOLOR="#FFFFFF">

                     <TD BGCOLOR="#c0c0c0">Name</TD>
 
                     <TD BGCOLOR="#c0c0c0">Details</TD>

                     <TD BGCOLOR="#c0c0c0">Status</TD>
 
                     <TD BGCOLOR="#c0c0c0">Time</TD>

                  </TR>

                  <xsl:choose>
                  
                     <xsl:when test="Report/Doc/DIter/Action/Step">
                        <xsl:apply-templates select="Report/Doc/DIter/Action/Step" />
                     </xsl:when>
                     
                     <xsl:when test="Report/Doc/Action/Step">
                        <xsl:apply-templates select="Report/Doc/Action/Step" />
                     </xsl:when>                                     
                    
                     
                  </xsl:choose>
                   <!--xsl:apply-templates select="Report/Doc/Action/Step" /-->
               </TABLE>

               <BR />
            </CENTER>

            <CENTER>
               <h3>
                  <FONT FACE="Times New Roman">Execution Completed : 
                  <xsl:value-of select="//Report/Doc/Summary/@eTime" />
                  </FONT>
               </h3>
            </CENTER>
         </BODY>
      </HTML>
   </xsl:template>

   <xsl:template match="Step">
      <xsl:if test="NodeArgs/@eType='User' or NodeArgs/@status = 'Failed' or NodeArgs/@status = 'Warnings' ">
         <TR>
            <TD class="hl0">
               <xsl:value-of select="NodeArgs/Disp" />
            </TD>

            <TD class="hl0">
               <xsl:value-of select="Details" disable-output-escaping="yes" />
            </TD>
            <TD class="hl0">
               <xsl:choose>
                  <xsl:when test="NodeArgs/@status = 'Passed'">
                     <FONT COLOR="#009900">
                        <xsl:value-of select="NodeArgs/@status" />
                     </FONT>
                  </xsl:when>

                  <xsl:when test="NodeArgs/@status = 'Failed'">
                     <FONT COLOR="Red">
                        <xsl:value-of select="NodeArgs/@status" />
                     </FONT>
                  </xsl:when>
               </xsl:choose>
            </TD>

            <TD class="hl0">
               <xsl:value-of select="Time" />
            </TD>

         </TR>
      </xsl:if>

      <xsl:apply-templates select="*[@rID]" />
   </xsl:template>
</xsl:stylesheet>

