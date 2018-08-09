 package com.myga.pdf;

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

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class DEMONSTRATION_HYPOTHETICAL_MVA {
	/*
	 * public static ArrayList<Integer> ActresultSet; public static
	 * ArrayList<Integer> ExpresultSet;
	 */

	/*static String ExpResultsFile = "D:\\Sagicor\\MYGA\\SLIC_MYGA_Expected_Results_Single_Page_v28_2018-06-26.xlsx";
	static String ActResultsFile = "D:\\Sagicor\\Myga\\Myga_ActualResult5.xlsx";
	static String ExpSheetName = "myga003";
	static String TextFilepath = "D:\\Sagicor\\Myga\\PDFToText_myga001.txt";
	static String ActSheetName = "myga003";
	static String pdfFilePath = "D:\\Sagicor\\Myga\\myga003.pdf";*/

	static String restLineValue2;
	static String restLineValue1;
	static String restLineValue;
	static String[] splitDataSet;
	static String[] splitDataSet1;
	static String[] splitDataSet2;
	static String FindValue;
	static String TerminateValue;

	/*public static void main(String[] args) throws Exception {
		ExtentReports extent = new ExtentReports("D:\\mFolder\\testMyga.html");

		ExtentTest testInst = extent.startTest("test with testcomplte");

		output_30_FIAValidation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
				pdfFilePath);
		extent.endTest(testInst);
		extent.flush();
	}*/

	public static void output_30_FIAValidation(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath)
			throws Exception, IOException {
		pdftoText(pdfFilePath, TextFilepath);

		MVADemoFIA14_ExcelReader(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
				pdfFilePath);

		CompareExcels(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName, pdfFilePath);

	}


	public static String CompareExcels(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) {
		List<List<Integer>> Actarray = new ArrayList<List<Integer>>();
		List<List<Integer>> Exparray = new ArrayList<List<Integer>>();
		Xlsx_Reader ExpResults = new Xlsx_Reader(ExpResultsFile);
		Xlsx_Reader ActResults = new Xlsx_Reader(ActResultsFile);
		try {
			// HashMap<Integer, String[]> results = new HashMap<Integer, String[]>();
			int counter = 1;

			for (int i = 3; i <= ActResults.getRowCount(ActSheetName); i++) {
				for (int j = 0; j < 9; j++) {
					String Actdata = ActResults.getCellFormulaData(ActSheetName, j + 12, i);
					String Expdata = ExpResults.getCellFormulaData(ExpSheetName, j + 11, i + 1);
					if (Actdata.equals(Expdata)) {
						testInst.log(LogStatus.PASS, Actdata + "actual value from " + ActSheetName + " matching with " +  Expdata
								+ " expected value from expected sheet" +  ExpSheetName);
					} else {
						List<Integer> ActresultSet = new ArrayList<Integer>();
						List<Integer> ExpresultSet = new ArrayList<Integer>();
						Actarray.add(ActresultSet);
						Exparray.add(ExpresultSet);
						ActresultSet.add(j + 12);
						ActresultSet.add(i);
						ExpresultSet.add(j + 11);
						ExpresultSet.add(i + 1);
						// ActResults.setCellColor(ActSheetName, j+11, i, "FAIL");
						// ExpResults.setCellColor(ExpSheetName, j+11, i+1, "FAIL");
						testInst.log(LogStatus.FAIL,
								Actdata + "actual value from " + ActSheetName + "does not match with " + Expdata
										+ " expected value from expected sheet" + ExpSheetName);
					}
					// return "PASS";
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		}
		System.out.println("File compare is Done");
		if (Actarray.size() != 0) {
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

	// When Decleard rate and Global is zero
	public static String MVADemoFIA14_ExcelReader(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {
		ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
		ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
		ArrayList<String> setCellList_Str = new ArrayList<String>();
		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
		try {
			StringBuilder sb = new StringBuilder();
			String line1 = br.readLine();
			String line = br.readLine();
			int lineNumber = 1;
			int rowNumber = 3;

			for (int i = 0; i <= lineNumber; i++) {

				if (Character.isDigit(line.charAt(0))) {
					// System.out.println("Line==" + line);
					String[] splitDataSet = line.split("\\s+");
					if (splitDataSet.length == 10 && (splitDataSet[0].length() == 1 || splitDataSet[0].length() == 2)
							&& (splitDataSet[2].length() > 4)) {

						for (int j = 0; j < splitDataSet.length; j++) {
							setCellList_intColumn.add(11 + j);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(splitDataSet[j]);
							// setCellList_Str.add("");
							/*
							 * PDFResults.setCellData(ActSheetName, 11 + 0, rowNumber, splitDataSet[1]);
							 * PDFResults.setCellData(ActSheetName, 11 + 1, rowNumber, splitDataSet[2]);
							 * PDFResults.setCellData(ActSheetName, 11 + 2, rowNumber, splitDataSet[3]);
							 * PDFResults.setCellData(ActSheetName, 11 + 3, rowNumber, splitDataSet[4]);
							 * PDFResults.setCellData(ActSheetName, 11 + 4, rowNumber, splitDataSet[5]);
							 * PDFResults.setCellData(ActSheetName, 11 + 5, rowNumber, splitDataSet[6]);
							 * PDFResults.setCellData(ActSheetName, 11 + 6, rowNumber, splitDataSet[7]);
							 * PDFResults.setCellData(ActSheetName, 11 + 7, rowNumber, splitDataSet[8]);
							 * PDFResults.setCellData(ActSheetName, 11 + 8, rowNumber, splitDataSet[9]);
							 */
							if (splitDataSet.length == j + 1) {
								rowNumber++;
							}

							lineNumber++;
						}

					}
					sb.append(line);
					sb.append(System.lineSeparator());
				}

				line = br.readLine();
				lineNumber++;
			}
			String everything = sb.toString();
			return "PASS";
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		} finally {
			PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
			System.out.println("Excel reading end");
			br.close();
			

		}
	}


}
