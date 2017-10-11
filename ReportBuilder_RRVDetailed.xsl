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
              TIMESTAMP
            </th>
          </tr>
          <xsl:for-each select="Report/Doc/Action/Step">
            <tr style="height:20px">
              <td align="left">
                <xsl:value-of select="Obj"/>
              </td>
              <td align="left">
                <xsl:value-of select="Details" disable-output-escaping="yes"/>
              </td>
              <td align="center">
                <xsl:choose>
                  <xsl:when test="NodeArgs/@status='Error'">
                    <font color="red">
                      ERROR
                    </font>
                  </xsl:when>
                  <xsl:when test="NodeArgs/@status='Failed'">
                    <font color="red">
                      FAILED
                    </font>
                  </xsl:when>
                  <xsl:when test="NodeArgs/@status='Passed'">
                    <font color="green">
                      PASSED
                    </font>
                  </xsl:when>
                  <xsl:when test="NodeArgs/@status='Warning'">
                    <font color="orange">
                      WARNING
                    </font>
                  </xsl:when>
                  <xsl:otherwise>
                        <font color="black">
                          <xsl:value-of select="NodeArgs/@status"/>
                        </font>
                  </xsl:otherwise>
                </xsl:choose>
                <td align="center">
                  <xsl:value-of select="Time" />
                </td>
              </td>
            </tr>
          </xsl:for-each>
        </table>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>