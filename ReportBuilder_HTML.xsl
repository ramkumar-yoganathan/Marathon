<?xml version='1.0'?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/">
    <html>
      <head>
        <style type="text/css">
          @import url(http://fonts.googleapis.com/css?family=Open+Sans:700);
          * {
          font-family: 'Open Sans', sans-serif !important;
          font-size: 10pt;
          }
          table {
          width: 80%;
          vertical-align: top;
          font-weight: 500;
          }
          html, body {
          padding: 0px !important;
          margin: 0px !important;
          height: 99%;
          min-height: 99%;
          }
          tr:nth-child(even) {background-color: #E3F2FD}
        </style>
      </head>
      <body>
          <table width="80%" align="center">
            <tr style="background-color:#149AC6;">
              <th style="color: white;">
                STEP NAME
              </th>
              <th style="color: white;">
                DESCRIPTION
              </th>
              <th style="color: white;">
                STATUS
              </th>
              <th style="color: white;">
                TIME STAMP
              </th>
            </tr>
            <xsl:for-each select="/Results/ReportNode/ReportNode/ReportNode/Data">
              <xsl:if test="Result = 'Failed' or Result = 'Warnings' or Result = 'Passed' or Result = 'Error'">
                <tr style="height:20px">
                  <td align="left">
                    <xsl:value-of select="Name"/>
                  </td>
                  <td align="left">
                    <xsl:value-of select="Description" disable-output-escaping="yes"/>
                  </td>
                  <td align="center">
                    <xsl:choose>
                      <xsl:when test="Result='Failed'">
                        <font color="red">
                          FAILED
                        </font>
                      </xsl:when>
                      <xsl:when test="Result='Passed'">
                        <font color="green">
                          PASSED
                        </font>
                      </xsl:when>
                      <xsl:when test="Result='Warning'">
                        <font color="orange">
                          WARNING
                        </font>
                      </xsl:when>
                    </xsl:choose>
                  </td>
                  <td align="center">
                    <xsl:value-of select="StartTime" />
                  </td>
                </tr>
              </xsl:if>
            </xsl:for-each>
          </table>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>