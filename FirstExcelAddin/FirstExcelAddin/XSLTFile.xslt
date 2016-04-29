<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
>
  <xsl:output method="xml" indent="yes"/>

  <xsl:template name="string-replace-all">
    <xsl:param name="text" />
    <xsl:param name="replace" />
    <xsl:param name="by" />
    <xsl:choose>
      <xsl:when test="contains($text, $replace)">
        <xsl:value-of select="substring-before($text,$replace)" />
        <xsl:value-of select="$by" />
        <xsl:call-template name="string-replace-all">
          <xsl:with-param name="text" select="substring-after($text,$replace)" />
          <xsl:with-param name="replace" select="$replace" />
          <xsl:with-param name="by" select="$by" />
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$text" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="/">
    <html>
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
        <script src="Scripts\jquery-1.9.1.min.js">
          <xsl:text> </xsl:text>
        </script>
        <script src="Scripts\bootstrap.min.js">
          <xsl:text> </xsl:text>
        </script>
        <script src="Scripts\Manipulation.js">
          <xsl:text> </xsl:text>
        </script>
        <link rel="stylesheet" type="text/css" href="Content\bootstrap.min.css" />
        <link rel="stylesheet" type="text/css" href="Content\test.css" />



      </head>
      <body title="Spreadsheet">


        <xsl:for-each select="Spreadsheet">
          <div class="page-header">
            <h1 class="text-center">
              <xsl:value-of select="@name"/>
            </h1>
          </div>
          <div class="container">
            <div class="col-md-12 col-lg-12">
              <div class="panel panel-default">
                <div class="panel-heading">
                  <h3 class="panel-title">
                    <b>Spreadsheet Description</b>
                  </h3>
                </div>
                <div class="panel-body">
                  <xsl:value-of select="@description" />
                  <hr/>
                  <div class="panel panel-default"  style="margin-bottom:0px;">
                    <div class="panel-heading">
                      <h3 class="panel-title">
                        <b>Worksheets</b>
                      </h3>
                    </div>
                    <div class="panel-body">
                      <div class="tabbable">
                        <ul class="nav nav-tabs worksheetTab">

                          <xsl:for-each select="Worksheet">
                            <xsl:variable name="worksheetName" select="@name" />
                            <xsl:choose>
                              <xsl:when test="position( )=1">
                                <li role="navigation" class="active">
                                  <a href="#{$worksheetName}" data-toggle="tab">
                                    <xsl:value-of select="$worksheetName" />
                                  </a>
                                </li>
                              </xsl:when>
                              <xsl:otherwise>
                                <li role="navigation">
                                  <a href="#{$worksheetName}" data-toggle="tab">
                                    <xsl:value-of select="$worksheetName" />
                                  </a>
                                </li>
                              </xsl:otherwise>
                            </xsl:choose>
                          </xsl:for-each>
                        </ul>
                        <div class="tab-content">
                          <xsl:for-each select="Worksheet">
                            <xsl:variable name="elementTotal" select="count(*)"/>
                            <xsl:variable name="cellTotal" select="count(Cell)"/>
                            <xsl:variable name="inputTotal" select="count(Input)"/>
                            <xsl:variable name="outputTotal" select="count(Output)"/>
                            <xsl:variable name="rangeTotal" select="count(Range)"/>
                            <xsl:variable name="columnTotal" select="count(Column)"/>
                            <xsl:variable name="rowTotal" select="count(Row)"/>
                            <xsl:choose>
                              <xsl:when test="position( )=1">
                                <div class="tab-pane active" id="{@name}">
                                  <div class="panel panel-default" style="margin-bottom:0px;">
                                    <div class="panel-heading">
                                      <h3 class="panel-title selectedItemTabHeader" />
                                    </div>
                                    <div class="panel-body">
                                      <xsl:value-of select="@description"/>
                                    </div>
                                  </div>
                                  <xsl:if test="$elementTotal > 0">
                                    <div class="panel panel-default" >
                                      <div class="panel-heading">
                                        <h3 class="panel-title selectedItemTab"/>
                                      </div>
                                      <xsl:if test="$cellTotal > 0">
                                        <div class="panel-body" style="padding-bottom:0px;">
                                          <div class="panel panel-default"  style="margin-bottom:0px;">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Cells</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Cell'">
                                                      <li>
                                                        <a href=".{@name}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Cell'">
                                                      <div class="tab-pane {@name}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Cell description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                        <xsl:if test="*">
                                                          <xsl:for-each select="*">
                                                            <div class="panel panel-default" style="margin-bottom:0px;">
                                                              <div class="panel-heading">
                                                                <h3 class="panel-title">
                                                                  <b>
                                                                    <xsl:value-of select="name()" />
                                                                    <xsl:text> </xsl:text>
                                                                    <xsl:value-of select="@cell" />
                                                                  </b>
                                                                </h3>
                                                              </div>
                                                              <div class="panel-body">
                                                                <xsl:value-of select="@description" />
                                                              </div>
                                                            </div>
                                                          </xsl:for-each>
                                                        </xsl:if>
                                                      </div>

                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                      <xsl:if test="$inputTotal > 0">
                                        <div class="panel-body"  style="padding-top:0px; padding-bottom:0px">
                                          <div class="panel panel-default" style="margin-bottom:0px;">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Inputs</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs teste">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Input'">
                                                      <li>
                                                        <a href=".{@name}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Input'">
                                                      <div class="tab-pane {@name}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Input description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                      <xsl:if test="$outputTotal > 0">
                                        <div class="panel-body"  style="padding-top:0px; padding-bottom:0px">
                                          <div class="panel panel-default" style="margin-bottom:0px;">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Outputs</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Output'">
                                                      <li>
                                                        <a href=".{@name}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Output'">
                                                      <div class="tab-pane {@name}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Output description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                      <xsl:if test="$columnTotal > 0">
                                        <div class="panel-body"  style="padding-top:0px; padding-bottom:0px">
                                          <div class="panel panel-default" style="margin-bottom:0px;">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Columns</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Column'">
                                                      <li>
                                                        <a href=".{@name}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Column'">
                                                      <div class="tab-pane {@name}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Column description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                      <xsl:if test="$rowTotal > 0">
                                        <div class="panel-body"  style="padding-top:0px; padding-bottom:0px;">
                                          <div class="panel panel-default" style="margin-bottom:0px;">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Rows</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Row'">
                                                      <li>
                                                        <a href=".{@name}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Row'">
                                                      <div class="tab-pane {@name}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Row description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                      <xsl:if test="$rangeTotal > 0">
                                        <div class="panel-body"  style="padding-top:0px; padding-bottom:0px;">
                                          <div class="panel panel-default">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Ranges</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Range'">
                                                      <xsl:variable name="changeChar">
                                                        <xsl:call-template name="string-replace-all">
                                                          <xsl:with-param name="text" select="@name" />
                                                          <xsl:with-param name="replace" select="':'" />
                                                          <xsl:with-param name="by" select="'-'" />
                                                        </xsl:call-template>
                                                      </xsl:variable>
                                                      <li>
                                                        <a href=".{$changeChar}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Range'">
                                                      <xsl:variable name="changeChar">
                                                        <xsl:call-template name="string-replace-all">
                                                          <xsl:with-param name="text" select="@name" />
                                                          <xsl:with-param name="replace" select="':'" />
                                                          <xsl:with-param name="by" select="'-'" />
                                                        </xsl:call-template>
                                                      </xsl:variable>
                                                      <div class="tab-pane {$changeChar}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Range description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                    </div>
                                  </xsl:if>
                                </div>
                              </xsl:when>
                              <xsl:otherwise>
                                <div class="tab-pane" id="{@name}">
                                  <div class="panel panel-default" style="margin-bottom:0px;">
                                    <div class="panel-heading">
                                      <h3 class="panel-title selectedItemTabHeader" />
                                    </div>
                                    <div class="panel-body">
                                      <xsl:value-of select="@description"/>
                                    </div>
                                  </div>
                                  <xsl:if test="$elementTotal > 0">
                                    <div class="panel panel-default" >
                                      <div class="panel-heading">
                                        <h3 class="panel-title selectedItemTab"/>
                                      </div>
                                      <xsl:if test="$cellTotal > 0">
                                        <div class="panel-body" style="padding-bottom:0px;">
                                          <div class="panel panel-default"  style="margin-bottom:0px;">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Cells</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Cell'">
                                                      <li>
                                                        <a href=".{@name}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Cell'">
                                                      <div class="tab-pane {@name}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Cell description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                        <xsl:if test="*">
                                                          <xsl:for-each select="*">
                                                            <div class="panel panel-default" style="margin-bottom:0px;">
                                                              <div class="panel-heading">
                                                                <h3 class="panel-title">
                                                                  <b>
                                                                    <xsl:value-of select="name()" />
                                                                    <xsl:text> </xsl:text>
                                                                    <xsl:value-of select="@cell" />
                                                                  </b>
                                                                </h3>
                                                              </div>
                                                              <div class="panel-body">
                                                                <xsl:value-of select="@description" />
                                                              </div>
                                                            </div>
                                                          </xsl:for-each>
                                                        </xsl:if>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                      <xsl:if test="$inputTotal > 0">
                                        <div class="panel-body"  style="padding-top:0px; padding-bottom:0px">
                                          <div class="panel panel-default" style="margin-bottom:0px;">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Inputs</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Input'">
                                                      <li>
                                                        <a href=".{@name}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Input'">
                                                      <div class="tab-pane {@name}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Input description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                      <xsl:if test="$outputTotal > 0">
                                        <div class="panel-body"  style="padding-top:0px; padding-bottom:0px">
                                          <div class="panel panel-default" style="margin-bottom:0px;">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Outputs</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Output'">
                                                      <li>
                                                        <a href=".{@name}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Output'">
                                                      <div class="tab-pane {@name}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Output description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                      <xsl:if test="$columnTotal > 0">
                                        <div class="panel-body"  style="padding-top:0px; padding-bottom:0px">
                                          <div class="panel panel-default" style="margin-bottom:0px;">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Columns</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Column'">
                                                      <li>
                                                        <a href=".{@name}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Column'">
                                                      <div class="tab-pane {@name}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Column description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                      <xsl:if test="$rowTotal > 0">
                                        <div class="panel-body"  style="padding-top:0px; padding-bottom:0px;">
                                          <div class="panel panel-default" style="margin-bottom:0px;">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Rows</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Row'">
                                                      <li>
                                                        <a href=".{@name}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Row'">
                                                      <div class="tab-pane {@name}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Row description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                      <xsl:if test="$rangeTotal > 0">
                                        <div class="panel-body"  style="padding-top:0px; padding-bottom:0px;">
                                          <div class="panel panel-default">
                                            <div class="panel-heading">
                                              <h3 class="panel-title">
                                                <b>Ranges</b>
                                              </h3>
                                            </div>
                                            <div class="panel-body">
                                              <div class="tabbable">
                                                <ul class="nav nav-tabs">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Range'">
                                                      <xsl:variable name="changeChar">
                                                        <xsl:call-template name="string-replace-all">
                                                          <xsl:with-param name="text" select="@name" />
                                                          <xsl:with-param name="replace" select="':'" />
                                                          <xsl:with-param name="by" select="'-'" />
                                                        </xsl:call-template>
                                                      </xsl:variable>
                                                      <li>
                                                        <a href=".{$changeChar}" data-toggle="tab">
                                                          <xsl:value-of select="@name" />
                                                        </a>
                                                      </li>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </ul>
                                                <div class="tab-content">
                                                  <xsl:for-each select="node()">
                                                    <xsl:variable name="currentNode" select="name()" />
                                                    <xsl:if test="$currentNode = 'Range'">
                                                      <xsl:variable name="changeChar">
                                                        <xsl:call-template name="string-replace-all">
                                                          <xsl:with-param name="text" select="@name" />
                                                          <xsl:with-param name="replace" select="':'" />
                                                          <xsl:with-param name="by" select="'-'" />
                                                        </xsl:call-template>
                                                      </xsl:variable>
                                                      <div class="tab-pane {$changeChar}">
                                                        <div class="panel panel-default" style="margin-bottom:0px;">
                                                          <div class="panel-heading">
                                                            <h3 class="panel-title">
                                                              <b>Range description</b>
                                                            </h3>
                                                          </div>
                                                          <div class="panel-body">
                                                            <xsl:value-of select="@description" />
                                                          </div>
                                                        </div>
                                                      </div>
                                                    </xsl:if>
                                                  </xsl:for-each>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      </xsl:if>
                                    </div>
                                  </xsl:if>
                                </div>
                              </xsl:otherwise>
                            </xsl:choose>
                          </xsl:for-each>
                        </div>
                      </div>
                    </div>
                  </div>
                    </div>
                  </div>
                </div>
              </div>
            </xsl:for-each>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>
