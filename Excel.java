/**
 * 
 */
package com.tsystems.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.record.chart.BarRecord;
import org.apache.poi.hssf.usermodel.HSSFChart;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartData;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.ScatterChartData;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author mariusc
 *
 */
public class Excel {

	
	public static  void fillValues(String fileName) {
		 //
        // Create an instance of workbook and sheet
        //		
		XSSFWorkbook workbook = new XSSFWorkbook();
        //HSSFWorkbook workbook = new HSSFWorkbook();
        //HSSFSheet sheet = workbook.createSheet();
		XSSFSheet sheet = workbook.createSheet();
 
        //
        // Create an instance of HSSFCellStyle which will be use to format the
        // cell. Here we define the cell top and bottom border and we also
        // define the background color.
        //
        XSSFCellStyle style = workbook.createCellStyle();
        style.setBorderTop((short) 6); // double lines border
        style.setBorderLeft((short) 2);
        style.setBorderBottom((short) 1); // single line border        
 
        
        //
        // We also define the font that we are going to use for displaying the
        // data of the cell. We set the font to ARIAL with 20pt in size and
        // make it BOLD and give blue as the color.
        //
        XSSFFont font = workbook.createFont();
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setFontHeightInPoints((short) 20);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setColor(HSSFColor.RED.index);
        style.setFont(font);
 
        //
        // We create a simple cell, set its value and apply the cell style.
        //
        XSSFRow row = sheet.createRow(1);
        XSSFCell cell = row.createCell(1);
        cell.setCellValue(new XSSFRichTextString("Hello Marius! This is a test message"));
        cell.setCellStyle(style);       
        sheet.autoSizeColumn((short) 1);
        
        XSSFRow row2 = sheet.createRow(2);
        XSSFCellStyle style2 = workbook.createCellStyle();
        style2.setFillBackgroundColor(HSSFColor.GOLD.index);
        XSSFCell cell2 = row2.createCell(1);
        cell2.setCellValue(new XSSFRichTextString("Hello all"));
        XSSFFont font2 = workbook.createFont();
        font2.setFontName(HSSFFont.FONT_ARIAL);
        font2.setFontHeightInPoints((short) 10);
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font2.setColor(HSSFColor.GREEN.index);
        style2.setFont(font2);
        cell2.setCellStyle(style2);       
        cell2.setCellType(XSSFCell.CELL_TYPE_BOOLEAN);        
       // sheet.autoSizeColumn((short) 1);

 
        //
        // Finally we write out the workbook into an excel file.
        //
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(new File(fileName));
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fos != null) {
                try {
                    fos.flush();
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
	}
	
	
	public static void readData(String filename) throws IOException  {
		 //
        // Create an ArrayList to store the data read from excel sheet.
        //
        List sheetData = new ArrayList();
 
        FileInputStream fis = null;
        try {
            //
            // Create a FileInputStream that will be use to read the
            // excel file.
            //
            fis = new FileInputStream(filename);
 
            //
            // Create an excel workbook from the file system.
            //
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            //
            // Get the first sheet on the workbook.
            //
            XSSFSheet sheet = workbook.getSheetAt(0);
 
            //
            // When we have a sheet object in hand we can iterator on
            // each sheet's rows and on each row's cells. We store the
            // data read on an ArrayList so that we can printed the
            // content of the excel to the console.
            //
            Iterator rows = sheet.rowIterator();
            while (rows.hasNext()) {
                XSSFRow row = (XSSFRow) rows.next();
                Iterator cells = row.cellIterator();
 
                List data = new ArrayList();
                while (cells.hasNext()) {
                    XSSFCell cell = (XSSFCell) cells.next();
                    if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
		              cell.setCellValue("bubulache");	                	
	                }
                    data.add(cell);
                }
 
                sheetData.add(data);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                fis.close();
            }
        }
 
        showExcelData(sheetData);
	}
	
	 private static void showExcelData(List sheetData) {
	        //
	        // Iterates the data and print it out to the console.
	        //
	        for (int i = 0; i < sheetData.size(); i++) {
	            List list = (List) sheetData.get(i);
	            for (int j = 0; j < list.size(); j++) {
	                XSSFCell cell = (XSSFCell) list.get(j);
	                if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
		                System.out.print(
		                        cell.getRichStringCellValue().getString());
	                	
	                }
	                if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
		                System.out.print(
		                        cell.getNumericCellValue());
	                	
	                }

	                
	                if (j < list.size() - 1) {
	                    System.out.print(", ");
	                }
	            }
	            System.out.println("");
	        }
	    }
	 
	 public static void addScatterChart( String fileName) throws Exception {
		 Workbook wb = new XSSFWorkbook();
	        Sheet sheet = wb.createSheet("Chart");
	        int NUM_OF_ROWS = 3;
	        int NUM_OF_COLUMNS = 10;
	        for(int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++)
	        {
	            Row row = sheet.createRow((short)rowIndex);
	            for(int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++)
	            {
	                Cell cell = row.createCell((short)colIndex);
	                cell.setCellValue(colIndex * (rowIndex + 1));
	            }

	        }

	        Drawing drawing = sheet.createDrawingPatriarch();
	        org.apache.poi.ss.usermodel.ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 15);
	        Chart chart = drawing.createChart(anchor);
	        ChartLegend legend = chart.getOrCreateLegend();
	        legend.setPosition(LegendPosition.LEFT);
	        ScatterChartData data = chart.getChartDataFactory().createScatterChartData();
	        ValueAxis bottomAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.BOTTOM);
	        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
	        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
	        org.apache.poi.ss.usermodel.charts.ChartDataSource xs = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, 9));
	        org.apache.poi.ss.usermodel.charts.ChartDataSource ys1 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 9));
	        org.apache.poi.ss.usermodel.charts.ChartDataSource ys2 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, 9));
	        data.addSerie(xs, ys1);
	        data.addSerie(xs, ys2);
	        chart.plot(data, new ChartAxis[] {
	            bottomAxis, leftAxis
	        });
	        FileOutputStream fileOut = new FileOutputStream(fileName);
	        wb.write(fileOut);
	        fileOut.close();
	 }
	 
	 public static void createBarChart(String fileName) throws Exception {
		 BarRecord rec = new BarRecord();

		 
		 Workbook wb = new XSSFWorkbook();
	        Sheet sheet = wb.createSheet("Chart");
	        int NUM_OF_ROWS = 3;
	        int NUM_OF_COLUMNS = 10;
	        for(int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++)
	        {
	            Row row = sheet.createRow((short)rowIndex);
	            for(int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++)
	            {
	                Cell cell = row.createCell((short)colIndex);
	                cell.setCellValue(colIndex * (rowIndex + 1));
	            }

	        }

	        Drawing drawing = sheet.createDrawingPatriarch();
	        org.apache.poi.ss.usermodel.ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 15);
	        Chart chart = drawing.createChart(anchor);
	        ChartLegend legend = chart.getOrCreateLegend();
	        legend.setPosition(LegendPosition.LEFT);
	        ScatterChartData data = chart.getChartDataFactory().createScatterChartData();
	        ValueAxis bottomAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.BOTTOM);
	        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
	        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
	        org.apache.poi.ss.usermodel.charts.ChartDataSource xs = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, 9));
	        org.apache.poi.ss.usermodel.charts.ChartDataSource ys1 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 9));
	        org.apache.poi.ss.usermodel.charts.ChartDataSource ys2 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, 9));
	        data.addSerie(xs, ys1);
	        data.addSerie(xs, ys2);
	        chart.plot(data, new ChartAxis[] {
	            bottomAxis, leftAxis
	        });
	        	        
	        FileOutputStream fileOut = new FileOutputStream(fileName);
	        wb.write(fileOut);
	        fileOut.close();
	 }
	 
	 
	 public static void update(String fileName) throws Exception {
		 //
	        // Create an ArrayList to store the data read from excel sheet.
	        //
	        List sheetData = new ArrayList();
	 
	        FileInputStream fis = null;
	        try {
	            //
	            // Create a FileInputStream that will be use to read the
	            // excel file.
	            //
	            fis = new FileInputStream(fileName);
	 
	            //
	            // Create an excel workbook from the file system.
	            //
	            XSSFWorkbook workbook = new XSSFWorkbook(fis);
	            //
	            // Get the first sheet on the workbook.
	            //
	            XSSFSheet sheet = workbook.getSheetAt(0);
	 
	            //
	            // When we have a sheet object in hand we can iterator on
	            // each sheet's rows and on each row's cells. We store the
	            // data read on an ArrayList so that we can printed the
	            // content of the excel to the console.
	            //
	            Iterator rows = sheet.rowIterator();
	            while (rows.hasNext()) {
	                XSSFRow row = (XSSFRow) rows.next();
	                XSSFCell cell = row.getCell(0);	                
	                if (cell == null) {
	                	row.createCell(0);
	                	cell = row.getCell(0);	   
	                	cell.setCellValue(1);
	                } else {
	                	cell.setCellValue(2);	
	                }
	                
//	                Iterator cells = row.cellIterator();
//	 
//	                List data = new ArrayList();
//	                while (cells.hasNext()) {
//	                    //XSSFCell cell = (XSSFCell) cells.next();
//	                    if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
//			              cell.setCellValue(row.getRowNum());	                	
//		                }
//	                    data.add(cell);
//	                }
//	 
//	                sheetData.add(data);
	            }
	            //
	            // Finally we write out the workbook into an excel file.
	            //
	            FileOutputStream fos =  new FileOutputStream(new File(fileName));
	                workbook.write(fos);
	           
	        } catch (IOException e) {
	            e.printStackTrace();
	        } finally {
	            if (fis != null) {
	                fis.close();
	            }
	        }
	 }
}
