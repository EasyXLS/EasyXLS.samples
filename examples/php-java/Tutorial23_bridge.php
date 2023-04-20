<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*================================================================
	 | Tutorial 23
	 |
	 | This tutorial shows how to create an Excel file in PHP with a 
	 | chart and how to set chart type and formatting properties for 
	 | chart area, plot area, axis, series and legend.
	  ==============================================================*/
	
	include("Format.inc");
	include("Color.inc");
	include("Chart.inc");
	include("LineStyleFormat.inc");
	include("ShadowFormat.inc");

	header("Content-Type: text/html");
	
	echo "Tutorial 23<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Create a worksheet
	$workbook->easy_addWorksheet("SourceData");
	
	// Get the table of data for the worksheet
	$xlsTable1 = $workbook->easy_getSheet("SourceData")->easy_getExcelTable();

	// Add data in cells for report header
	$xlsTable1->easy_getCell(0, 0)->setValue("Show Date");
	$xlsTable1->easy_getCell(0, 1)->setValue("Available Places");
	$xlsTable1->easy_getCell(0, 2)->setValue("Available Tickets");
	$xlsTable1->easy_getCell(0, 3)->setValue("Sold Tickets");

	// Add data in cells for chart report values
	$xlsTable1->easy_getCell(1, 0)->setValue("03/13/2005 00:00:00");
	$xlsTable1->easy_getCell(1, 0)->setFormat($FORMAT_FORMAT_DATE);
	$xlsTable1->easy_getCell(2, 0)->setValue("03/14/2005 00:00:00");
	$xlsTable1->easy_getCell(2, 0)->setFormat($FORMAT_FORMAT_DATE);
	$xlsTable1->easy_getCell(3, 0)->setValue("03/15/2005 00:00:00");
	$xlsTable1->easy_getCell(3, 0)->setFormat($FORMAT_FORMAT_DATE);
	$xlsTable1->easy_getCell(4, 0)->setValue("03/16/2005 00:00:00");
	$xlsTable1->easy_getCell(4, 0)->setFormat($FORMAT_FORMAT_DATE);

	$xlsTable1->easy_getCell(1, 1)->setValue("10000");
	$xlsTable1->easy_getCell(2, 1)->setValue("5000");
	$xlsTable1->easy_getCell(3, 1)->setValue("8500");
	$xlsTable1->easy_getCell(4, 1)->setValue("1000");

	$xlsTable1->easy_getCell(1, 2)->setValue("8000");
	$xlsTable1->easy_getCell(2, 2)->setValue("4000");
	$xlsTable1->easy_getCell(3, 2)->setValue("6000");
	$xlsTable1->easy_getCell(4, 2)->setValue("1000");

	$xlsTable1->easy_getCell(1, 3)->setValue("920");
	$xlsTable1->easy_getCell(2, 3)->setValue("1005");
	$xlsTable1->easy_getCell(3, 3)->setValue("342");
	$xlsTable1->easy_getCell(4, 3)->setValue("967");

	// Set column widths
	$xlsTable1->easy_getColumnAt(0)->setWidth(100);
	$xlsTable1->easy_getColumnAt(1)->setWidth(100);
	$xlsTable1->easy_getColumnAt(2)->setWidth(100);
	$xlsTable1->easy_getColumnAt(3)->setWidth(100);

	// Add a chart sheet
	echo "";
	$workbook->easy_addChart("Chart", "=SourceData!\$A$1:\$D$5", $CHART_SERIES_IN_COLUMNS);
	
	// Get the previously added chart
	$xlsChartSheet = $workbook->easy_getSheetAt(1);
	$xlsChart = $xlsChartSheet->easy_getExcelChart();

	// Set chart type
	$xlsChart->easy_setChartType($CHART_CHART_TYPE_CYLINDER_COLUMN);

	// Format chart area
	$xlsChartArea = $xlsChart->easy_getChartArea();
	$xlsChartArea->getLineColorFormat()->setLineColor(java("java.awt.Color")->DARKGRAY);
	$xlsChartArea->getLineStyleFormat()->setDashType($LINESTYLEFORMAT_DASH_TYPE_SOLID);
	$xlsChartArea->getLineStyleFormat()->setWidth(0.25);
	
	// Format chart plot area
	$xlsPlotArea = $xlsChart->easy_getPlotArea();
	$xlsPlotArea->getLineColorFormat()->setLineColor(java("java.awt.Color")->DARKGRAY);
	$xlsPlotArea->getLineStyleFormat()->setDashType($LINESTYLEFORMAT_DASH_TYPE_SOLID);
	$xlsPlotArea->getLineStyleFormat()->setWidth(0.25);

	// Format chart legend
	$xlsChatLegend = $xlsChart->easy_getLegend();
	$xlsChatLegend->getFillFormat()->setBackground(java("java.awt.Color")->PINK);
	$xlsChatLegend->getFontFormat()->setForeground(java("java.awt.Color")->BLUE);
	$xlsChatLegend->getFontFormat()->setItalic(true);
	$xlsChatLegend->setKeysArrangementDirection($CHART_KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL);
	$xlsChatLegend->setPlacement($CHART_LEGEND_CORNER);
	$xlsChatLegend->getShadowFormat()->setShadow($SHADOWFORMAT_OFFSET_DIAGONAL_BOTTOM_RIGHT);

	// Format chart X axis
	$xlsXAxis = $xlsChart->easy_getCategoryXAxis();
	$xlsXAxis->getLineColorFormat()->setLineColor(java("java.awt.Color")->LIGHTGRAY);
	$xlsXAxis->getLineStyleFormat()->setDashType($LINESTYLEFORMAT_DASH_TYPE_DASH_DOT);
	$xlsXAxis->getLineStyleFormat()->setWidth(0.25);
	$xlsXAxis->getFontFormat()->setForeground(java("java.awt.Color")->RED);

	// Format chart Y axis
	$xlsYAxis = $xlsChart->easy_getValueYAxis();
	$xlsYAxis->getLineColorFormat()->setLineColor(java("java.awt.Color")->LIGHTGRAY);
	$xlsYAxis->getLineStyleFormat()->setDashType($LINESTYLEFORMAT_DASH_TYPE_LONG_DASH);
	$xlsYAxis->getLineStyleFormat()->setWidth(0.25);
	$xlsYAxis->getFontFormat()->setForeground(java("java.awt.Color")->BLUE);

	// Fomat chart series
	$xlsChart->easy_getSeriesAt(0)->getFillFormat()->setBackground(java("java.awt.Color")->BLUE);
	$xlsChart->easy_getSeriesAt(1)->getFillFormat()->setBackground(java("java.awt.Color")->YELLOW);
	$xlsChart->easy_getSeriesAt(2)->getFillFormat()->setBackground(java("java.awt.Color")->GREEN);
	
	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial23 - various Excel chart settings.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial23 - various Excel chart settings.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>