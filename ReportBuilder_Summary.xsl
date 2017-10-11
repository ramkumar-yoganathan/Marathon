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
					<!--tr:nth-child(even) {background-color: #E3F2FD} -->
				</style>
			</head>
			<body>
				<br></br>
				<table width="100%">
					<table width="100%" style="" border="1" align="center">
						<xsl:for-each select="/Environment/Summary">
							<tr style="height:25px;">
								<td bgcolor="A7B6DA">
									HOST NAME
								</td>
								<td align="center">
									<xsl:value-of select="HostName"/>
								</td>
								<td bgcolor="A7B6DA">
									START TIME
								</td>
								<td align="center">
									<xsl:value-of select="Starttime"/>
								</td>
							</tr>
							<tr style="height:25px;">
								<td bgcolor="A7B6DA">
									TIME ZONE
								</td>
								<td align="center">
									<xsl:value-of select="Timezone"/>
								</td>
								<td bgcolor="A7B6DA">
									END TIME
								</td>
								<td align="center">
									<xsl:value-of select="Endtime"/>
								</td>
							</tr>
							<tr style="height:25px;">
								<td bgcolor="A7B6DA">
									BROWSER
								</td>
								<td align="center">
									<xsl:value-of select="Browser"/>
								</td>
								<td bgcolor="A7B6DA">
									TIME ELAPSED
								</td>
								<td align="center">
									<xsl:value-of select="Elapsed"/>
								</td>
							</tr>
							<tr style="height:25px;">
								<td bgcolor="A7B6DA">
									USER NAME
								</td>
								<td align="center">
									<xsl:value-of select="User"/>
								</td>
								<td bgcolor="A7B6DA">
									PRODUCT
								</td>
								<td align="center">
									<xsl:value-of select="Version"/>
								</td>
							</tr>
						</xsl:for-each>
					</table>
					<table width="80%" border="1" align="center">
						<tr style="background-color:#149AC6;font-size:12">
							<th style="color: white;">
								S.NO
							</th>
							<th style="color: white; width: 30%;">
								TEST NAME
							</th>
							<th style="color: white;">
								STATUS
							</th>
							<th style="color: white;">
								DURATION
							</th>
							<th style="color: white;">
								STEP PASSED
							</th>
							<th style="color: white;">
								STEP FAILED
							</th>
							<th style="color: white;">
								STEP WARNINGS
							</th>
							<th style="color: white;">
								CUSTOM REPORT
							</th>
							<th style="color: white;">
								UFT REPORT
							</th>
							<th style="color: white;">
								SESSION VIDEO
							</th>
						</tr>
						<xsl:for-each select="/Environment/Test">
							<tr style="height:25px">
								<td align="center" style="width:10px">
									<xsl:value-of select="Sequence"/>
								</td>
								<td align="center" style="width:500px; padding: 3px;text-align:left">
									<xsl:value-of select="Name"/>
								</td>
								<td align="center" width="10%">
									<xsl:choose>
										<xsl:when test="Status='FAILED'">
											<xsl:attribute name="style">background-color:#bb0000</xsl:attribute>
										</xsl:when>
										<xsl:when test="Status='PASSED'">
											<xsl:attribute name="style">background-color:#00bb00</xsl:attribute>
										</xsl:when>
										<xsl:when test="Status='WARNING'">
											<xsl:attribute name="style">background-color:#FFFF00</xsl:attribute>
										</xsl:when>
										<xsl:when test="Status='STOPPED'">
											<xsl:attribute name="style">background-color:##483D8B</xsl:attribute>
										</xsl:when>
										<xsl:otherwise>
											<xsl:attribute name="style">background-color:#c0c0c0</xsl:attribute>
										</xsl:otherwise>
									</xsl:choose>
									<xsl:value-of select="Status"/>
								</td>
								<td align="center" width="20%">
									<xsl:value-of select="Duration"/>
								</td>
								<td align="center" width="20%">
									<xsl:attribute name="style">foreground-color:#00bb00</xsl:attribute>
									<xsl:value-of select="StepPassed"/>
								</td>
								<td align="center" width="20%">
									<xsl:attribute name="style">foreground-color:#bb0000</xsl:attribute>
									<xsl:value-of select="StepFailed"/>
								</td>
								<td align="center" width="20%">
									<xsl:attribute name="style">foreground-color:#FFFF00</xsl:attribute>
									<xsl:value-of select="StepWarning"/>
								</td>
								<td align="center" width="20%">
									<xsl:choose>
										<xsl:when test="Report != ''">
											<a>
												<xsl:attribute name="href">
													<xsl:value-of select="StepReport" />
												</xsl:attribute>REPORT
											</a>
										</xsl:when>
										<xsl:otherwise>
											<xsl:attribute name="href">
												<xsl:value-of select="StepReport" />
											</xsl:attribute>
										</xsl:otherwise>
									</xsl:choose>
								</td>
								<td align="center" width="40%">
									<xsl:choose>
										<xsl:when test="Report != ''">
											<a>
												<xsl:attribute name="href">
													<xsl:value-of select="Report" />
												</xsl:attribute>REPORT
											</a>
										</xsl:when>
										<xsl:otherwise>
											<xsl:attribute name="href">
												<xsl:value-of select="Report" />
											</xsl:attribute>
										</xsl:otherwise>
									</xsl:choose>
								</td>
								<td align="center" width="40%">
									<xsl:choose>
										<xsl:when test="Movie != ''">
											<a>
												<xsl:attribute name="href">
													<xsl:value-of select="Movie" />
												</xsl:attribute>VIDEO
											</a>
										</xsl:when>
										<xsl:otherwise>
											<xsl:attribute name="href">
												<xsl:value-of select="Movie" />
											</xsl:attribute>
										</xsl:otherwise>
									</xsl:choose>
								</td>
							</tr>
						</xsl:for-each>
					</table>
				</table>
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>