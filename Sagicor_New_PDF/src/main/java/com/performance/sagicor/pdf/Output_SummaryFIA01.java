package com.performance.sagicor.pdf;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import com.myga.pdf.Xlsx_Reader;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class Output_SummaryFIA01 {

/*	static String ExpResultsFile = "D:\\Sagicor_New_Final\\SLIC_SEC_Expected_Results_Single_Page_v29_2018-07-19 - Copy.xlsx";
	static String ActResultsFile = "D:\\Sagicor_New_Final\\NewActualresultSagicor_New.xlsx";
	static String ExpSheetName = "SEC014";
	static String TextFilepath= "D:\\Sagicor_New_Final\\PDFToText_SEC001.txt";
	static String ActSheetName = "SEC001";
	static String pdfFilePath= "D:\\Sagicor_New_Final\\SEC014.pdf";*/

	static String restLineValue2;
	static String restLineValue1;
	static String restLineValue;
	static String[] splitDataSet;
	
	
	

	/*public static void main(String[] args) throws Exception {
		ExtentReports extent = new ExtentReports("D:\\mFolder\\test1.html");

		ExtentTest testInst = extent.startTest("test with testcomplte");

		Output_SummaryFIA01validation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath,
				ActSheetName, pdfFilePath);
		extent.endTest(testInst);
		extent.flush();
	}*/

	public static void Output_SummaryFIA01validation(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {
		pdftoText(pdfFilePath, TextFilepath);
		String result = GetStrategy(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
				pdfFilePath, "Declared Rate Strategy", 6);
		String DRS = result.split("&")[0];
		String SNPS = result.split("&")[1];
		String GMIS = result.split("&")[2];


		if (!DRS.startsWith("0%") && !SNPS.equals("0%") && !GMIS.equals("0%")) {
			result = Output_SummaryFIA14_readExcel(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);

		} else if (DRS.equals("0%") && !SNPS.equals("0%") && GMIS.equals("0%")) {

			result = DecleredRateAndGlobalIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath,
					ActSheetName, pdfFilePath);
		} else if (DRS.equals("0%") && SNPS.equals("0%") && !GMIS.equals("0%")) {

			result = DecleredRateAndSNPIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath,
					ActSheetName, pdfFilePath);
		} else if (!DRS.equals("0%") && SNPS.equals("0%") && GMIS.equals("0%")) {

			result = SNPANDGlobalIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath,
					ActSheetName, pdfFilePath);
		} else if (DRS.equals("0%") && !SNPS.equals("0%") && !GMIS.equals("0%")) {

			result = DecleredIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);
		} else if (!DRS.equals("0%") && !SNPS.equals("0%") && GMIS.equals("0%")) {

			result = GlobalIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);
		} else if (!DRS.equals("0%") && SNPS.equals("0%") && !GMIS.equals("0%")) {
			result = SNPIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);
		}

		 CompareExcels_Output_SummaryFIA01validation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
				TextFilepath, ActSheetName, pdfFilePath);

		//RecordFailResults(testInst,  ExpResultsFile, ActResultsFile,results);
	}

	public static String GetStrategy(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath, String findValue,
			int columnNumber) throws Exception {
		String DeclaredRateStrategy = null;
		String SNP500IndexStrategy = null;
		String GlobalMultiIndexStrategy = null;
		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
		try {
			StringBuilder sb = new StringBuilder();
			// String line1 = br.readLine();
			String line = br.readLine();
			int lineNumber = 1;
			int rowNumber = 3;

			for (int i = 0; i <= lineNumber; i++) {
				// System.out.println("lineNumber==" + lineNumber);

				if (line.contains("Declared Rate Strategy")) {
					restLineValue = line.replaceAll("Declared Rate Strategy", " ");
					restLineValue = restLineValue.trim();
					splitDataSet = restLineValue.split("\\s+");
					DeclaredRateStrategy = splitDataSet[0];
				} else if (line.contains("S&P 500® Index Strategy")) {
					restLineValue = line.replaceAll("S&P 500® Index Strategy", " ");
					restLineValue = restLineValue.trim();
					splitDataSet = restLineValue.split("\\s+");
					SNP500IndexStrategy = splitDataSet[0];
				} else if (line.contains("Global Multi-Index Strategy")) {
					restLineValue = line.replaceAll("Global Multi-Index Strategy", " ");
					restLineValue = restLineValue.trim();
					splitDataSet = restLineValue.split("\\s+");
					GlobalMultiIndexStrategy = splitDataSet[0];
				}
				line = br.readLine();
				lineNumber++;
			}
		} catch (Exception e) {
			e.printStackTrace();
			return DeclaredRateStrategy + "&" + SNP500IndexStrategy + "&" + GlobalMultiIndexStrategy;
		}
		return DeclaredRateStrategy + "&" + SNP500IndexStrategy + "&" + GlobalMultiIndexStrategy;
	}
	public static String Output_SummaryFIA14_readExcel(ExtentTest testInst, String ExpResultsFile,
			String ActResultsFile, String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath)
			throws Exception {

		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
		ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
		ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
		ArrayList<String> setCellList_Str = new ArrayList<String>();
		try {
			StringBuilder sb = new StringBuilder();
			// String line1 = br.readLine();
			String line = br.readLine();
			int lineNumber = 1;
			int rowNumber = 3;

			for (int i = 0; i <= lineNumber; i++) {
				// System.out.println("lineNumber==" + lineNumber);
				if (Character.isDigit(line.charAt(0))) {
					//System.out.println("Line==" + line);
					String[] splitDataSet = line.split("\\s+");
					// System.out.println("splitData Length=" + splitDataSet.length);

					for (int j = 0; j < splitDataSet.length; j++) {
						if ((splitDataSet.length==13)&&(splitDataSet[0].length() == 1|| splitDataSet[0].length() == 2)
								&& (splitDataSet[1].length() == 2) ) {
							String data1 = splitDataSet[j];
							// setCellData(String sheetName,int colName,int rowNum, String data)
							setCellList_intColumn.add(j+63);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(data1);
							//PDFResults.setCellData(ActSheetName, j + 69, rowNumber, data1);
							if (splitDataSet.length == j + 1) {
								rowNumber++;
							}
						}
					}
					sb.append(line);
					sb.append(System.lineSeparator());
				}
				line = br.readLine();
				lineNumber++;
			}
			String everything = sb.toString();
			// System.out.println(everything);
			return "PASS";
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		} finally {
			PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
			System.out.println("text to excel is Done");
			br.close();
		}
	}
	
	
	// When Decleard rate and Global is zero
		public static String DecleredRateAndGlobalIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
				String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

			Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
			BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
			ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
			ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
			ArrayList<String> setCellList_Str = new ArrayList<String>();
			try {
				StringBuilder sb = new StringBuilder();
				String line1 = br.readLine();
				String line = br.readLine();
				int lineNumber = 1;
				int rowNumber = 3;

				for (int i = 0; i <= lineNumber; i++) {
					// System.out.println("lineNumber==" + lineNumber);
					if (Character.isDigit(line.charAt(0))) {
						// System.out.println("Line==" + line);
						String[] splitDataSet = line.split("\\s+");
						// System.out.println("splitData Length=" + splitDataSet.length);
						// String data1 = splitDataSet[j];

						// for (int j = 0; j < splitDataSet.length; j++) {
						if (splitDataSet[0].length() == 10 && splitDataSet[1].length() == 2
								&& splitDataSet[2].length() == 2) {
							// String data1 = splitDataSet[j];
							
							setCellList_intColumn.add( 69 + 0);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[0]);
							setCellList_intColumn.add( 69 + 1);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[1]);
							setCellList_intColumn.add( 69 + 2);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[2]);
							setCellList_intColumn.add( 69 + 3);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[3]);
							setCellList_intColumn.add( 69 + 4);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[4]);
							setCellList_intColumn.add( 69 + 5);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[5]);
							setCellList_intColumn.add( 69 + 6);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[6]);
							setCellList_intColumn.add( 69 + 7);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add("0");
							setCellList_intColumn.add( 69 + 8);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[7]);
							setCellList_intColumn.add( 69 + 9);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add("0");
							setCellList_intColumn.add( 69 + 10);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[8]);
							setCellList_intColumn.add( 69 + 11);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[9]);
							setCellList_intColumn.add( 69 + 12);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[10]);
							setCellList_intColumn.add( 69 + 13);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[11]);
							
						/*	PDFResults.setCellData(ActSheetName, 69 + 0, rowNumber, splitDataSet[0]);
							PDFResults.setCellData(ActSheetName, 69 + 1, rowNumber, splitDataSet[1]);
							PDFResults.setCellData(ActSheetName, 69 + 2, rowNumber, splitDataSet[2]);
							PDFResults.setCellData(ActSheetName, 69 + 3, rowNumber, splitDataSet[3]);
							PDFResults.setCellData(ActSheetName, 69 + 4, rowNumber, splitDataSet[4]);
							PDFResults.setCellData(ActSheetName, 69 + 5, rowNumber, splitDataSet[5]);
							PDFResults.setCellData(ActSheetName, 69 + 6, rowNumber, splitDataSet[6]);
							PDFResults.setCellData(ActSheetName, 69 + 7, rowNumber, "0");
							PDFResults.setCellData(ActSheetName, 69 + 8, rowNumber, splitDataSet[7]);
							PDFResults.setCellData(ActSheetName, 69 + 9, rowNumber, "0");
							PDFResults.setCellData(ActSheetName, 69 + 10, rowNumber, splitDataSet[8]);
							PDFResults.setCellData(ActSheetName, 69 + 11, rowNumber, splitDataSet[9]);
							PDFResults.setCellData(ActSheetName, 69 + 12, rowNumber, splitDataSet[10]);
							PDFResults.setCellData(ActSheetName, 69 + 13, rowNumber, splitDataSet[11]);*/
							rowNumber++;

						}
						// }
						sb.append(line);
						sb.append(System.lineSeparator());
					}
					line = br.readLine();
					lineNumber++;
				}
				String everything = sb.toString();
				// System.out.println(everything);
				return "PASS";
			} catch (Exception e) {
				e.printStackTrace();
				return "FAIL";
			} finally {
				PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
				System.out.println("text to excel is Done");
				br.close();
			}
		}
		
		
		
		// When Decleared rate and SnP is zero and Global is not zero
		public static String DecleredRateAndSNPIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
				String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

			Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
			BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
			ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
			ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
			ArrayList<String> setCellList_Str = new ArrayList<String>();
			try {
				StringBuilder sb = new StringBuilder();
				String line1 = br.readLine();
				String line = br.readLine();
				int lineNumber = 1;
				int rowNumber = 3;

				for (int i = 0; i <= lineNumber; i++) {
					// System.out.println("lineNumber==" + lineNumber);
					if (Character.isDigit(line.charAt(0))) {
						// System.out.println("Line==" + line);
						String[] splitDataSet = line.split("\\s+");
						// System.out.println("splitData Length=" + splitDataSet.length);

						/// for (int j = 0; j < splitDataSet.length; j++) {
						if (splitDataSet[0].length() == 10 && splitDataSet[1].length() == 2
								&& splitDataSet[2].length() == 2) {
							// System.out.println(data1);
							// setCellData(String sheetName,int colName,int rowNum, String data)
							
							
							setCellList_intColumn.add( 69 + 0);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[0]);
							setCellList_intColumn.add( 69 + 1);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[1]);
							setCellList_intColumn.add( 69 + 2);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[2]);
							setCellList_intColumn.add( 69 + 3);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[3]);
							setCellList_intColumn.add( 69 + 4);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[4]);
							setCellList_intColumn.add( 69 + 5);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[5]);
							setCellList_intColumn.add( 69 + 6);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[6]);
							setCellList_intColumn.add( 69 + 7);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add("0");
							setCellList_intColumn.add( 69 + 8);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add("0");
							setCellList_intColumn.add( 69 + 9);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[7]);
							setCellList_intColumn.add( 69 + 10);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[8]);
							setCellList_intColumn.add( 69 + 11);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[9]);
							setCellList_intColumn.add( 69 + 12);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[10]);
							setCellList_intColumn.add( 69 + 13);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[11]);

							/*PDFResults.setCellData(ActSheetName, 69 + 0, rowNumber, splitDataSet[0]);
							PDFResults.setCellData(ActSheetName, 69 + 1, rowNumber, splitDataSet[1]);
							PDFResults.setCellData(ActSheetName, 69 + 2, rowNumber, splitDataSet[2]);
							PDFResults.setCellData(ActSheetName, 69 + 3, rowNumber, splitDataSet[3]);
							PDFResults.setCellData(ActSheetName, 69 + 4, rowNumber, splitDataSet[4]);
							PDFResults.setCellData(ActSheetName, 69 + 5, rowNumber, splitDataSet[5]);
							PDFResults.setCellData(ActSheetName, 69 + 6, rowNumber, splitDataSet[6]);
							PDFResults.setCellData(ActSheetName, 69 + 7, rowNumber, "0");
							PDFResults.setCellData(ActSheetName, 69 + 8, rowNumber, "0");
							PDFResults.setCellData(ActSheetName, 69 + 9, rowNumber, splitDataSet[7]);
							PDFResults.setCellData(ActSheetName, 69 + 10, rowNumber, splitDataSet[8]);
							PDFResults.setCellData(ActSheetName, 69 + 11, rowNumber, splitDataSet[9]);
							PDFResults.setCellData(ActSheetName, 69 + 12, rowNumber, splitDataSet[10]);
							PDFResults.setCellData(ActSheetName, 69 + 13, rowNumber, splitDataSet[11]);*/
							rowNumber++;
						}
						// }
						sb.append(line);
						sb.append(System.lineSeparator());
					}
					line = br.readLine();
					lineNumber++;
				}
				String everything = sb.toString();
				// System.out.println(everything);
				return "PASS";
			} catch (Exception e) {
				e.printStackTrace();
				return "FAIL";
			} finally {
				PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
				System.out.println("text to excel is Done");
				br.close();
			}
		}

		// When Decleared rate is not zero and S&P,Global is zero
		public static String SNPANDGlobalIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
				String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

			Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
			BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
			ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
			ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
			ArrayList<String> setCellList_Str = new ArrayList<String>();
			try {
				StringBuilder sb = new StringBuilder();
				String line1 = br.readLine();
				String line = br.readLine();
				int lineNumber = 1;
				int rowNumber = 3;

				for (int i = 0; i <= lineNumber; i++) {
					// System.out.println("lineNumber==" + lineNumber);
					if (Character.isDigit(line.charAt(0))) {
						// System.out.println("Line==" + line);
						String[] splitDataSet = line.split("\\s+");
						// System.out.println("splitData Length=" + splitDataSet.length);

						/* for (int j = 0; j < splitDataSet.length; j++) { */
						if ((splitDataSet.length==11) && (splitDataSet[0].length() == 1||splitDataSet[0].length() == 2) && (splitDataSet[1].length() == 2)) {
							// System.out.println(data1);
							// setCellData(String sheetName,int colName,int rowNum, String data)
							
							setCellList_intColumn.add( 63 + 0);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[0]);
							setCellList_intColumn.add( 63 + 1);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[1]);
							setCellList_intColumn.add( 63 + 2);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[2]);
							setCellList_intColumn.add( 63 + 3);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[3]);
							setCellList_intColumn.add( 63 + 4);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[4]);
							setCellList_intColumn.add( 63 + 5);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[5]);
							setCellList_intColumn.add( 63 + 6);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[6]);
							setCellList_intColumn.add( 63 + 7);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add("0");
							setCellList_intColumn.add( 63 + 8);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add("0");
							setCellList_intColumn.add( 63 + 9);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[7]);
							setCellList_intColumn.add( 63 + 10);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[8]);
							setCellList_intColumn.add( 63 + 11);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[9]);
							setCellList_intColumn.add( 63 + 12);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[10]);
							/*setCellList_intColumn.add( 69 + 13);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[11]);*/


							/*PDFResults.setCellData(ActSheetName, 69 + 0, rowNumber, splitDataSet[0]);
							PDFResults.setCellData(ActSheetName, 69 + 1, rowNumber, splitDataSet[1]);
							PDFResults.setCellData(ActSheetName, 69 + 2, rowNumber, splitDataSet[2]);
							PDFResults.setCellData(ActSheetName, 69 + 3, rowNumber, splitDataSet[3]);
							PDFResults.setCellData(ActSheetName, 69 + 4, rowNumber, splitDataSet[4]);
							PDFResults.setCellData(ActSheetName, 69 + 5, rowNumber, splitDataSet[5]);
							PDFResults.setCellData(ActSheetName, 69 + 6, rowNumber, splitDataSet[6]);
							PDFResults.setCellData(ActSheetName, 69 + 7, rowNumber, splitDataSet[7]);
							PDFResults.setCellData(ActSheetName, 69 + 8, rowNumber, "0");
							PDFResults.setCellData(ActSheetName, 69 + 9, rowNumber, "0");
							PDFResults.setCellData(ActSheetName, 69 + 10, rowNumber, splitDataSet[8]);
							PDFResults.setCellData(ActSheetName, 69 + 11, rowNumber, splitDataSet[9]);
							PDFResults.setCellData(ActSheetName, 69 + 12, rowNumber, splitDataSet[10]);
							PDFResults.setCellData(ActSheetName, 69 + 13, rowNumber, splitDataSet[11]);*/
							rowNumber++;
						}
						sb.append(line);
						sb.append(System.lineSeparator());
					}
					line = br.readLine();
					lineNumber++;
				}
				String everything = sb.toString();
				// System.out.println(everything);
				return "PASS";
			} catch (Exception e) {
				e.printStackTrace();
				return "FAIL";
			} finally {
				PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
				System.out.println("text to excel is Done");
				br.close();
			}
		}
		
		
		
		// When Decleared rate is zero and S&P,Global is not zero
		public static String DecleredIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
				String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

			Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
			BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
			ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
			ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
			ArrayList<String> setCellList_Str = new ArrayList<String>();
			try {
				StringBuilder sb = new StringBuilder();
				String line1 = br.readLine();
				String line = br.readLine();
				int lineNumber = 1;
				int rowNumber = 3;

				for (int i = 0; i <= lineNumber; i++) {
					// System.out.println("lineNumber==" + lineNumber);
					if (Character.isDigit(line.charAt(0))) {
						// System.out.println("Line==" + line);
						String[] splitDataSet = line.split("\\s+");
						// System.out.println("splitData Length=" + splitDataSet.length);

						/* for (int j = 0; j < splitDataSet.length; j++) { */
						if (splitDataSet[0].length() == 10 && splitDataSet[1].length() == 2
								&& splitDataSet[2].length() == 2) {
							// System.out.println(data1);
							// setCellData(String sheetName,int colName,int rowNum, String data)
							
							setCellList_intColumn.add( 69 + 0);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[0]);
							setCellList_intColumn.add( 69 + 1);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[1]);
							setCellList_intColumn.add( 69 + 2);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[2]);
							setCellList_intColumn.add( 69 + 3);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[3]);
							setCellList_intColumn.add( 69 + 4);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[4]);
							setCellList_intColumn.add( 69 + 5);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[5]);
							setCellList_intColumn.add( 69 + 6);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[6]);
							setCellList_intColumn.add( 69 + 7);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add("0");
							setCellList_intColumn.add( 69 + 8);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[7]);
							setCellList_intColumn.add( 69 + 9);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[8]);
							setCellList_intColumn.add( 69 + 10);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[9]);
							setCellList_intColumn.add( 69 + 11);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[10]);
							setCellList_intColumn.add( 69 + 12);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[11]);
							setCellList_intColumn.add( 69 + 13);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[12]);


							/*PDFResults.setCellData(ActSheetName, 69 + 0, rowNumber, splitDataSet[0]);
							PDFResults.setCellData(ActSheetName, 69 + 1, rowNumber, splitDataSet[1]);
							PDFResults.setCellData(ActSheetName, 69 + 2, rowNumber, splitDataSet[2]);
							PDFResults.setCellData(ActSheetName, 69 + 3, rowNumber, splitDataSet[3]);
							PDFResults.setCellData(ActSheetName, 69 + 4, rowNumber, splitDataSet[4]);
							PDFResults.setCellData(ActSheetName, 69 + 5, rowNumber, splitDataSet[5]);
							PDFResults.setCellData(ActSheetName, 69 + 6, rowNumber, splitDataSet[6]);
							PDFResults.setCellData(ActSheetName, 69 + 7, rowNumber, "0");
							PDFResults.setCellData(ActSheetName, 69 + 8, rowNumber, splitDataSet[7]);
							PDFResults.setCellData(ActSheetName, 69 + 9, rowNumber, splitDataSet[8]);
							PDFResults.setCellData(ActSheetName, 69 + 10, rowNumber, splitDataSet[9]);
							PDFResults.setCellData(ActSheetName, 69 + 11, rowNumber, splitDataSet[10]);
							PDFResults.setCellData(ActSheetName, 69 + 12, rowNumber, splitDataSet[11]);
							PDFResults.setCellData(ActSheetName, 69 + 13, rowNumber, splitDataSet[12]);*/
							rowNumber++;
						}
						sb.append(line);
						sb.append(System.lineSeparator());
					}
					line = br.readLine();
					lineNumber++;
				}
				String everything = sb.toString();
				// System.out.println(everything);
				return "PASS";
			} catch (Exception e) {
				e.printStackTrace();
				return "FAIL";
			} finally {
				PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
				System.out.println("text to excel is Done");
				br.close();
			}
		}


		
		
		// When Decleared rate and S&P is not zero,Global is zero
		public static String GlobalIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
				String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

			Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
			BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
			ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
			ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
			ArrayList<String> setCellList_Str = new ArrayList<String>();
			try {
				StringBuilder sb = new StringBuilder();
				String line1 = br.readLine();
				String line = br.readLine();
				int lineNumber = 1;
				int rowNumber = 3;

				for (int i = 0; i <= lineNumber; i++) {
					// System.out.println("lineNumber==" + lineNumber);
					if (Character.isDigit(line.charAt(0))) {
						// System.out.println("Line==" + line);
						String[] splitDataSet = line.split("\\s+");
						// System.out.println("splitData Length=" + splitDataSet.length);

						/* for (int j = 0; j < splitDataSet.length; j++) { */
						if (splitDataSet[0].length() == 10 && splitDataSet[1].length() == 2
								&& splitDataSet[2].length() == 2) {
							// System.out.println(data1);
							// setCellData(String sheetName,int colName,int rowNum, String data)
							
							setCellList_intColumn.add( 69 + 0);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[0]);
							setCellList_intColumn.add( 69 + 1);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[1]);
							setCellList_intColumn.add( 69 + 2);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[2]);
							setCellList_intColumn.add( 69 + 3);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[3]);
							setCellList_intColumn.add( 69 + 4);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[4]);
							setCellList_intColumn.add( 69 + 5);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[5]);
							setCellList_intColumn.add( 69 + 6);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[6]);
							setCellList_intColumn.add( 69 + 7);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[7]);
							setCellList_intColumn.add( 69 + 8);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[8]);
							setCellList_intColumn.add( 69 + 9);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add("0");
							setCellList_intColumn.add( 69 + 10);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[9]);
							setCellList_intColumn.add( 69 + 11);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[10]);
							setCellList_intColumn.add( 69 + 12);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[11]);
							setCellList_intColumn.add( 69 + 13);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[12]);

							/*PDFResults.setCellData(ActSheetName, 69 + 0, rowNumber, splitDataSet[0]);
							PDFResults.setCellData(ActSheetName, 69 + 1, rowNumber, splitDataSet[1]);
							PDFResults.setCellData(ActSheetName, 69 + 2, rowNumber, splitDataSet[2]);
							PDFResults.setCellData(ActSheetName, 69 + 3, rowNumber, splitDataSet[3]);
							PDFResults.setCellData(ActSheetName, 69 + 4, rowNumber, splitDataSet[4]);
							PDFResults.setCellData(ActSheetName, 69 + 5, rowNumber, splitDataSet[5]);
							PDFResults.setCellData(ActSheetName, 69 + 6, rowNumber, splitDataSet[6]);
							PDFResults.setCellData(ActSheetName, 69 + 7, rowNumber, splitDataSet[7]);
							PDFResults.setCellData(ActSheetName, 69 + 8, rowNumber, splitDataSet[8]);
							PDFResults.setCellData(ActSheetName, 69 + 9, rowNumber, "0");
							PDFResults.setCellData(ActSheetName, 69 + 10, rowNumber, splitDataSet[9]);
							PDFResults.setCellData(ActSheetName, 69 + 11, rowNumber, splitDataSet[10]);
							PDFResults.setCellData(ActSheetName, 69 + 12, rowNumber, splitDataSet[11]);
							PDFResults.setCellData(ActSheetName, 69 + 13, rowNumber, splitDataSet[12]);*/
							rowNumber++;
						}
						sb.append(line);
						sb.append(System.lineSeparator());
					}
					line = br.readLine();
					lineNumber++;
				}
				String everything = sb.toString();
				// System.out.println(everything);
				return "PASS";
			} catch (Exception e) {
				e.printStackTrace();
				return "FAIL";
			} finally {
				PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
				System.out.println("text to excel is Done");
				br.close();
			}
		}
		
		
		
		// When Decleared rate and Global is not zero and S&P is zero.
		public static String SNPIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
				String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

			Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
			BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
			ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
			ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
			ArrayList<String> setCellList_Str = new ArrayList<String>();
			try {
				StringBuilder sb = new StringBuilder();
				String line1 = br.readLine();
				String line = br.readLine();
				int lineNumber = 1;
				int rowNumber = 3;

				for (int i = 0; i <= lineNumber; i++) {
					// System.out.println("lineNumber==" + lineNumber);
					if (Character.isDigit(line.charAt(0))) {
						// System.out.println("Line==" + line);
						String[] splitDataSet = line.split("\\s+");
						// System.out.println("splitData Length=" + splitDataSet.length);

						/* for (int j = 0; j < splitDataSet.length; j++) { */
						if (splitDataSet[0].length() == 10 && splitDataSet[1].length() == 2
								&& splitDataSet[2].length() == 2) {
							// System.out.println(data1);
							// setCellData(String sheetName,int colName,int rowNum, String data)
							
							
							setCellList_intColumn.add( 69 + 0);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[0]);
							setCellList_intColumn.add( 69 + 1);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[1]);
							setCellList_intColumn.add( 69 + 2);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[2]);
							setCellList_intColumn.add( 69 + 3);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[3]);
							setCellList_intColumn.add( 69 + 4);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[4]);
							setCellList_intColumn.add( 69 + 5);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[5]);
							setCellList_intColumn.add( 69 + 6);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[6]);
							setCellList_intColumn.add( 69 + 7);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[7]);
							setCellList_intColumn.add( 69 + 8);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add("0");
							setCellList_intColumn.add( 69 + 9);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[8]);
							setCellList_intColumn.add( 69 + 10);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[9]);
							setCellList_intColumn.add( 69 + 11);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[10]);
							setCellList_intColumn.add( 69 + 12);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[11]);
							setCellList_intColumn.add( 69 + 13);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[12]);

							/*PDFResults.setCellData(ActSheetName, 69 + 0, rowNumber, splitDataSet[0]);
							PDFResults.setCellData(ActSheetName, 69 + 1, rowNumber, splitDataSet[1]);
							PDFResults.setCellData(ActSheetName, 69 + 2, rowNumber, splitDataSet[2]);
							PDFResults.setCellData(ActSheetName, 69 + 3, rowNumber, splitDataSet[3]);
							PDFResults.setCellData(ActSheetName, 69 + 4, rowNumber, splitDataSet[4]);
							PDFResults.setCellData(ActSheetName, 69 + 5, rowNumber, splitDataSet[5]);
							PDFResults.setCellData(ActSheetName, 69 + 6, rowNumber, splitDataSet[6]);
							PDFResults.setCellData(ActSheetName, 69 + 7, rowNumber, splitDataSet[7]);
							PDFResults.setCellData(ActSheetName, 69 + 8, rowNumber, "0");
							PDFResults.setCellData(ActSheetName, 69 + 9, rowNumber, splitDataSet[8]);
							PDFResults.setCellData(ActSheetName, 69 + 10, rowNumber, splitDataSet[9]);
							PDFResults.setCellData(ActSheetName, 69 + 11, rowNumber, splitDataSet[10]);
							PDFResults.setCellData(ActSheetName, 69 + 12, rowNumber, splitDataSet[11]);
							PDFResults.setCellData(ActSheetName, 69 + 13, rowNumber, splitDataSet[12]);*/
							rowNumber++;
						}
						sb.append(line);
						sb.append(System.lineSeparator());
					}
					line = br.readLine();
					lineNumber++;
				}
				String everything = sb.toString();
				// System.out.println(everything);
				return "PASS";
			} catch (Exception e) {
				e.printStackTrace();
				return "FAIL";
			} finally {
				PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
				System.out.println("text to excel is Done");
				br.close();
			}
		}


	/* Method to put cell color for fail data */
		/*public static boolean RecordFailResults(ExtentTest testInst, String exp,String act,HashMap<Integer, String[]> results) {
			try {
				//System.out.println("results=="+results.size());
				
				Xlsx_Reader ExpResults = new Xlsx_Reader(exp);
				Xlsx_Reader ActResults = new Xlsx_Reader(act);

				for (Map.Entry<Integer, String[]> entry : results.entrySet()) {
					String[] actData = entry.getValue()[0].split("#");
					ActResults.setRedColor(actData[0], Integer.parseInt(actData[1]), Integer.parseInt(actData[2]));

					String[] expData = entry.getValue()[1].split("#");
					ExpResults.setRedColor(expData[0], Integer.parseInt(expData[1]), Integer.parseInt(expData[2]));
					testInst.log(LogStatus.FAIL, "Validation is failed at: column  sheet name: " + expData[0]
							+ "Actual result is : " + actData[3] + "Expected result is : " + expData[3]);
				}
				ActResults.writeAllData();
				ExpResults.writeAllData();
				return true;
			} catch (Exception e) {
				e.printStackTrace();
				return false;
			}
		}*/

	public static String CompareExcels_Output_SummaryFIA01validation(ExtentTest testInst, String ExpResultsFile,
			String ActResultsFile, String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) {
		List<List<Integer>> Actarray = new ArrayList<List<Integer>>();
		List<List<Integer>> Exparray = new ArrayList<List<Integer>>();
		Xlsx_Reader ExpResults = new Xlsx_Reader(ExpResultsFile);
		Xlsx_Reader ActResults = new Xlsx_Reader(ActResultsFile);
		try {
			//HashMap<Integer, String[]> results = new HashMap<Integer, String[]>();

			for (int i = 2; i <= ActResults.getRowCount(ActSheetName); i++) {
				for (int j = 1; j <= 14; j++) {
					String Actdata = ActResults.getCellFormulaData(ActSheetName, j + 62, i + 1);
					String Expdata = ExpResults.getCellFormulaData(ExpSheetName, j + 19, i + 3);
					if (Actdata.equals(Expdata)) {
						testInst.log(LogStatus.PASS, "Actual value " + Actdata + " from sheet " + ActSheetName
								+ "is matching with " + Expdata + "from expected sheet" + ExpSheetName);
					} else {
						List<Integer> ActresultSet = new ArrayList<Integer>();
						List<Integer> ExpresultSet = new ArrayList<Integer>();
						Actarray.add(ActresultSet);
						Exparray.add(ExpresultSet);
						ActresultSet.add(j+62);
						ActresultSet.add(i+1);
						ExpresultSet.add(j+19);
						ExpresultSet.add(i+3);
						//ActResults.setCellColor(ActSheetName, j+68, i+1, "FAIL");
						//ExpResults.setCellColor(ExpSheetName, j+20, i+3, "FAIL");
						testInst.log(LogStatus.FAIL, Actdata + "actual value from " + ActSheetName + "does not match with " + Expdata + " expected value from expected sheet" + ExpSheetName );
					}
					// return "PASS";
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		}
		System.out.println("File compare is Done");
		if(Actarray.size()!=0) {
			ActResults.setCellColor(ActSheetName, Actarray);
			ExpResults.setCellColor(ExpSheetName, Exparray);
		}
		return "PASS";

	}

	public static String pdftoText(String pdfFilePath, String TextFilepath) throws InterruptedException, IOException {
		// APP_LOGS.debug("Click on Button");
		try {
			PDFParser parser;
			String parsedText = null;
			PDFTextStripper pdfStripper = null;
			PDDocument pdDoc = null;
			COSDocument cosDoc = null;

			pdDoc = PDDocument.load(new File(pdfFilePath));
			pdfStripper = new PDFTextStripper();

			String content = pdfStripper.getText(pdDoc);

			File file = new File(TextFilepath);

			// if file doesnt exists, then create it
			if (!file.exists()) {
				file.createNewFile();
			}

			FileWriter fw = new FileWriter(file.getAbsoluteFile());
			BufferedWriter bw = new BufferedWriter(fw);
			bw.write(content);
			bw.close();

			System.out.println("Pdf to text Done");
			return "PASS";

		} catch (IOException e) {
			e.printStackTrace();
			return "FAIL";
		}
	}

}
