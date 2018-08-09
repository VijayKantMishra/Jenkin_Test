package com.myga.pdf;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class Xlsx_Reader {
	// public String path;
	public   FileInputStream fis = null;
	public  FileOutputStream fileOut = null;
	public  File excel;
	private  XSSFWorkbook workbook = null;
	private  XSSFSheet sheet = null;
	private  Row row = null;
	private  Cell cell = null;
	private  String path;

	public Xlsx_Reader(String path) {
		this.path = path;
		try {
			excel = new File(path);
			fis = new FileInputStream(excel);

			// fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	// returns the row count in a sheet
	public int getRowCount(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		sheet = workbook.getSheetAt(index);
		int number = sheet.getLastRowNum() + 1;
		return number;
	}

	// returns number of columns in a sheet
	public int getColumnCount(String sheetName) {
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);
		return row.getLastCellNum();
	}

	// returns the data from a cell
	public String getCellData(String sheetName, String colName, int rowNum) {
		try {
			int index = workbook.getSheetIndex(sheetName);
			int col_Num = 0;
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					col_Num = i;
			}
			// sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			cell = row.getCell(col_Num);
			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				return cell.getStringCellValue();
			} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				String cellText = String.valueOf(cell.getNumericCellValue());
				return cellText;
			} else {
				return "";
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return "";
		}
	}

	// find whether sheets exists
	public boolean isSheetExist(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			index = workbook.getSheetIndex(sheetName.toUpperCase());
			if (index == -1) {
				return false;
			} else {
				return true;
			}
		} else {
			return true;
		}
	}

	public boolean setCellData(String sheetName, String colName, int rowNum, String data) {
		try {
			int colNum = 0;
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);

			int index = workbook.getSheetIndex(sheetName);
			System.out.println("sheetName=" + sheetName + " " + colName + " " + rowNum + " " + data);
			sheet = workbook.getSheetAt(index);
			XSSFCellStyle style = workbook.createCellStyle();
			XSSFFont style1 = workbook.createFont();
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);
			cell.setCellValue(data);
			// set cell color
			if (data.equals("PASS")) {
				System.out.println("PASS");
				style.setFillForegroundColor(new XSSFColor(new java.awt.Color(34, 139, 34)));
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
				style1.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
				style.setFont(style1);
				cell.setCellStyle(style);
			} else if (data.equals("FAIL")) {
				System.out.println("FAIL");
				style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 0, 0)));
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
				style1.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
				style.setFont(style1);
				cell.setCellStyle(style);
			} else if (data.equals("NO RUN")) {
				System.out.println("NO RUN" + "34,139,34");
				style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 222, 173)));
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
				style1.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
				style.setFont(style1);
				cell.setCellStyle(style);
			} else if (data.contains("::::")) {
				style.setFillForegroundColor(new XSSFColor(new java.awt.Color(230, 230, 250)));
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
				style1.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
				style.setFont(style1);
				cell.setCellStyle(style);
			}
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}

	public boolean setCellData(String sheetName,int colNum,int rowNum, String data){
		  try{
		   //int colNum=0;
		   fis = new FileInputStream(path);
		   //File file = new File(path);
		  // file.setWritable(true);
		   
		   workbook = new XSSFWorkbook(fis);
		   
		   int index = workbook.getSheetIndex(sheetName);
		   //System.out.println("sheetName="+sheetName+" "+colNum+" "+rowNum+" "+data);
		   sheet = workbook.getSheetAt(index);
		   XSSFCellStyle style = workbook.createCellStyle();
		   XSSFFont style1 = workbook.createFont();
		   row=sheet.getRow(0);
/*		   for(int i=0;i<row.getLastCellNum();i++){
		    if(row.getCell(i).getStringCellValue().trim().equals(colName))
		     colNum=i;
		    }
*/		   sheet.autoSizeColumn(colNum); 
		   row = sheet.getRow(rowNum-1);
		   if (row == null)
		    row = sheet.createRow(rowNum-1);
		    cell = row.getCell(colNum); 
		   if (cell == null)
		    cell = row.createCell(colNum);
		   cell.setCellValue(data);
		   //set cell color
		   if(data.equals("PASS")){
		    System.out.println("PASS");
		    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(34,139,34)));
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
		    style1.setColor(new XSSFColor(new java.awt.Color(0,0,0)));
		    style.setFont(style1);
		    cell.setCellStyle(style);
		   }else if(data.equals("FAIL")){
		    System.out.println("FAIL");
		    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255,0,0)));
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
		    style1.setColor(new XSSFColor(new java.awt.Color(0,0,0)));
		    style.setFont(style1); 
		    cell.setCellStyle(style);
		   }else if(data.equals("NO RUN")){
		    System.out.println("NO RUN"+"34,139,34");
		    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255,222,173)));
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
		    style1.setColor(new XSSFColor(new java.awt.Color(0,0,0)));
		    style.setFont(style1); 
		    cell.setCellStyle(style);
		   }else if(data.contains("::::")){
		    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(230,230,250)));
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
		    style1.setColor(new XSSFColor(new java.awt.Color(0,0,0)));
		    style.setFont(style1);
		    cell.setCellStyle(style);
		   }
		   fileOut = new FileOutputStream(path);
		   workbook.write(fileOut);
		   fileOut.close(); 
		 //  file.setReadOnly();
		   return true;
		  }catch(Exception e){
		   e.printStackTrace();
		   return false;
		   }
		  }
	
	public boolean setCellData_Perform(String sheetName,List<Integer> listColumn, List<Integer> listRow,List<String> listDataArray){
		  try{
		   //int colNum=0;
		   fis = new FileInputStream(path);
		   //File file = new File(path);
		  // file.setWritable(true);
		   
		   workbook = new XSSFWorkbook(fis);
		   
		   int index = workbook.getSheetIndex(sheetName);
		   //System.out.println("sheetName="+sheetName+" "+data);
		   sheet = workbook.getSheetAt(index);
		   XSSFCellStyle style = workbook.createCellStyle();
		   XSSFFont style1 = workbook.createFont();
		   row=sheet.getRow(0);
/*		   for(int i=0;i<row.getLastCellNum();i++){
		    if(row.getCell(i).getStringCellValue().trim().equals(colName))
		     colNum=i;
		    }*/
		    for(int i=0; i<listDataArray.size(); i++) {
					int colNum=listColumn.get(i);
					int rowNum=listRow.get(i);
					String data =listDataArray.get(i);
					  //System.out.println("sheetName="+sheetName+" "+data);
		   sheet.autoSizeColumn(colNum); 
		   row = sheet.getRow(rowNum-1);
		   if (row == null)
		    row = sheet.createRow(rowNum-1);
		    cell = row.getCell(colNum); 
		   if (cell == null)
		    cell = row.createCell(colNum);
		   cell.setCellValue(data);
		   //set cell color
		   /*if(data.equals("PASS")){
		    System.out.println("PASS");
		    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(34,139,34)));
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
		    style1.setColor(new XSSFColor(new java.awt.Color(0,0,0)));
		    style.setFont(style1);
		    cell.setCellStyle(style);
		   }else if(data.equals("FAIL")){
		    System.out.println("FAIL");
		    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255,0,0)));
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
		    style1.setColor(new XSSFColor(new java.awt.Color(0,0,0)));
		    style.setFont(style1); 
		    cell.setCellStyle(style);
		   }else if(data.equals("NO RUN")){
		    System.out.println("NO RUN"+"34,139,34");
		    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255,222,173)));
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
		    style1.setColor(new XSSFColor(new java.awt.Color(0,0,0)));
		    style.setFont(style1); 
		    cell.setCellStyle(style);
		   }else if(data.contains("::::")){
		    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(230,230,250)));
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
		    style1.setColor(new XSSFColor(new java.awt.Color(0,0,0)));
		    style.setFont(style1);
		    cell.setCellStyle(style);
		    */	
		   }
		    fileOut = new FileOutputStream(path);
			   workbook.write(fileOut);
			   fileOut.close(); 
			   return true;
		  }
		  
	
		 //  file.setReadOnly();
		  
		  catch(Exception e){
		   e.printStackTrace();
		   return false;
		   }
		  }
	

	protected boolean setRedColor(String sheetName,int colNum,int rowNum) {
		try {
			int index = workbook.getSheetIndex(sheetName);
			//System.out.println("sheetName=" + sheetName + " " + colNum + " " + rowNum + " " );
			sheet = workbook.getSheetAt(index);
			XSSFCellStyle style = workbook.createCellStyle();
			XSSFFont style1 = workbook.createFont();
			row = sheet.getRow(0);

			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);

			//System.out.println("FAIL");
			style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 0, 0)));
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
			style1.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			style.setFont(style1);
			cell.setCellStyle(style);
			return true;
		}
		catch (Exception e) {
			e.printStackTrace();
			System.out.println("I got execption in set color");
			return false;
		}
	}
	
	public void writeAllData() {
		try {
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch (Exception e) {
			// TODO: handle exception
		}
	}



	public boolean setCellColor(String sheetName, int colNum, int rowNum, String data) {
		try {
			// int colNum=0;
			//fis = new FileInputStream(path);
			//workbook = new XSSFWorkbook(fis);

			int index = workbook.getSheetIndex(sheetName);
			System.out.println("sheetName=" + sheetName + " " + colNum + " " + rowNum + " " + data);
			sheet = workbook.getSheetAt(index);
			XSSFCellStyle style = workbook.createCellStyle();
			XSSFFont style1 = workbook.createFont();
			row = sheet.getRow(0);
			/*
			 * for(int i=0;i<row.getLastCellNum();i++){
			 * if(row.getCell(i).getStringCellValue().trim().equals(colName)) colNum=i; }
			 */
			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);
			// cell.setCellValue(data);
			// set cell color
			if (data.equals("PASS")) {
				System.out.println("PASS");
				style.setFillForegroundColor(new XSSFColor(new java.awt.Color(34, 139, 34)));
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
				style1.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
				style.setFont(style1);
				cell.setCellStyle(style);
			} else if (data.equals("FAIL")) {
				System.out.println("FAIL");
				style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 0, 0)));
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
				style1.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
				style.setFont(style1);
				cell.setCellStyle(style);
			} else if (data.equals("NO RUN")) {
				System.out.println("NO RUN" + "34,139,34");
				style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 222, 173)));
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
				style1.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
				style.setFont(style1);
				cell.setCellStyle(style);
			} else if (data.contains("::::")) {
				style.setFillForegroundColor(new XSSFColor(new java.awt.Color(230, 230, 250)));
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
				style1.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
				style.setFont(style1);
				cell.setCellStyle(style);
			}
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}
	
	
	public boolean setCellColor(String sheetName, List<List<Integer>> listArray) {
		try {
			// int colNum=0;
			//fis = new FileInputStream(path);
			//workbook = new XSSFWorkbook(fis);

			int index = workbook.getSheetIndex(sheetName);
			sheet = workbook.getSheetAt(index);
			XSSFCellStyle style = workbook.createCellStyle();
			XSSFFont style1 = workbook.createFont();
			row = sheet.getRow(0);
			/*
			 * for(int i=0;i<row.getLastCellNum();i++){
			 * if(row.getCell(i).getStringCellValue().trim().equals(colName)) colNum=i; }
			 */
			for(int i=0; i<listArray.size(); i++) {
					int colNum=listArray.get(i).get(0);
					int rowNum=listArray.get(i).get(1);
					
					sheet.autoSizeColumn(colNum);
					row = sheet.getRow(rowNum - 1);
					if (row == null)
						row = sheet.createRow(rowNum - 1);
					cell = row.getCell(colNum);
					if (cell == null)
						cell = row.createCell(colNum);
					// cell.setCellValue(data);
					// set cell color
					//System.out.println("FAIL");
					style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 0, 0)));
					style.setFillPattern(CellStyle.SOLID_FOREGROUND);
					style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
					style1.setColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
					style.setFont(style1);
					cell.setCellStyle(style);
			}
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}

	// returns the data from a cell
	public String getCellData(String sheetName, int colNum, int rowNum) {
		try {
			int index = workbook.getSheetIndex(sheetName);
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			cell = row.getCell(colNum);
			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				return cell.getStringCellValue();
			} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				System.out.println("cell.getNumericCellValue()=" + cell.getNumericCellValue());
				String cellText = String.valueOf(cell.getNumericCellValue());
				return cellText;
			} else {
				return "";
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "";
		}
	}

	// returns the data from a cell
	public String getCellFormulaData(String sheetName, int colNum, int rowNum) {
		try {
			//System.out.println(workbook.getNumberOfSheets());
			int index = workbook.getSheetIndex(sheetName);
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			cell = row.getCell(colNum);

			DataFormatter df = new DataFormatter();
			String value = df.formatCellValue(cell);
			return value;

			/*
			 * if(cell.getCellType()==Cell.CELL_TYPE_STRING){ return
			 * cell.getStringCellValue(); }else
			 * if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC ||
			 * cell.getCellType()==Cell.CELL_TYPE_FORMULA ){
			 * System.out.println("cell.getNumericCellValue()="+cell.getNumericCellValue());
			 * String cellText = String.valueOf(cell.getNumericCellValue()); return
			 * cellText; }else{ return ""; }
			 */
		} catch (Exception e) {
			e.printStackTrace();
			return "";
		}
	}

	// returns true if column is created successfully
	public boolean addColumn(String sheetName, String colName) {
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);

			XSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(new XSSFColor(new java.awt.Color(25, 25, 112)));
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);

			XSSFFont style1 = workbook.createFont();
			style1.setFontName(XSSFFont.DEFAULT_FONT_NAME);
			style1.setColor(new XSSFColor(new java.awt.Color(255, 255, 255)));
			style.setFont(style1);
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			if (row == null)
				row = sheet.createRow(0);
			if (row.getLastCellNum() == -1) {
				cell = row.createCell(0);
			} else {
				cell = row.createCell(row.getLastCellNum());
			}
			cell.setCellValue(colName);
			cell.setCellStyle(style);

			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}

	public boolean setCellData(String sheetName, String colName, int rowNum, int data) {
		try {
			int colNum = 0;
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);

			int index = workbook.getSheetIndex(sheetName);
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);
			cell.setCellValue(data);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}
}