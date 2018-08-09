package com.performance.sagicor.pdf;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import com.myga.pdf.Xlsx_Reader;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class Output_30_FIA01_NewActual {
	/*
	 * static String ExpResultsFile =
	 * "D:\\Sagicor_New_Final\\SLIC_SEC_Expected_Results_Single_Page_v29_2018-07-19 - Copy.xlsx"
	 * ; static String ActResultsFile =
	 * "D:\\Sagicor_New_Final\\NewActualresultSagicor_New.xlsx"; static String
	 * ExpSheetName = "SEC014"; static String TextFilepath=
	 * "D:\\Sagicor_New_Final\\PDFToText_SEC001.txt"; static String ActSheetName =
	 * "SEC001"; static String pdfFilePath= "D:\\Sagicor_New_Final\\SEC014.pdf";
	 */

	static String restLineValue2;
	static String restLineValue1;
	static String restLineValue;
	static String[] splitDataSet;
	static String[] splitDataSet1;
	static String[] splitDataSet2;
	static String FindValue;
	static String TerminateValue;

	/*
	 * public static void main(String[] args) throws Exception { ExtentReports
	 * extent = new ExtentReports("D:\\mFolder\\PAckagePerformance.html");
	 * 
	 * ExtentTest testInst = extent.startTest("test with testcomplte");
	 * 
	 * output_30_FIAValidation(testInst, ExpResultsFile, ActResultsFile,
	 * ExpSheetName, TextFilepath, ActSheetName, pdfFilePath);
	 * extent.endTest(testInst); extent.flush(); }
	 */
	public static void output_30_FIAValidation(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath)
			throws Exception, IOException {
		pdftoText(pdfFilePath, TextFilepath);
		PutPremium(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName, pdfFilePath);

		String result = GetStrategy(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
				pdfFilePath, "Declared Rate Strategy", 6);
		String DRS = result.split("&")[0];
		String SNPS = result.split("&")[1];
		String GMIS = result.split("&")[2];

		if (!DRS.startsWith("0%") && !SNPS.equals("0%") && !GMIS.equals("0%")) {
			result = ConvertToExcel(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
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

		CompareExcels(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName, pdfFilePath);

		// RecordFailResults(testInst, ExpResultsFile, ActResultsFile,results);

	}

	/* Method to put cell color for fail data */
	/*
	 * public static boolean RecordFailResults(ExtentTest testInst, String
	 * exp,String act,HashMap<Integer, String[]> results) { try {
	 * //System.out.println("results=="+results.size()); Xlsx_Reader ExpResults =
	 * new Xlsx_Reader(exp); Xlsx_Reader ActResults = new Xlsx_Reader(act);
	 * 
	 * for (Map.Entry<Integer, String[]> entry : results.entrySet()) { String[]
	 * actData = entry.getValue()[0].split("#"); ActResults.setRedColor(actData[0],
	 * Integer.parseInt(actData[1]), Integer.parseInt(actData[2]));
	 * 
	 * String[] expData = entry.getValue()[1].split("#");
	 * ExpResults.setRedColor(expData[0], Integer.parseInt(expData[1]),
	 * Integer.parseInt(expData[2])); testInst.log(LogStatus.FAIL,
	 * "Validation is failed at: column  sheet name: " + expData[0] +
	 * "Actual result is : " + actData[3] + "Expected result is : " + expData[3]); }
	 * ActResults.writeAllData(); ExpResults.writeAllData(); return true; } catch
	 * (Exception e) { e.printStackTrace(); return false; } }
	 */

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
				for (int j = 6; j < 19; j++) {
					String Actdata = ActResults.getCellFormulaData(ActSheetName, j, i);
					String Expdata = ExpResults.getCellFormulaData(ExpSheetName, j, i + 2);
					if (Actdata.equals(Expdata)) {
						testInst.log(LogStatus.PASS, "Actual value " + Actdata + " from sheet " + ActSheetName
								+ "is matching with " + Expdata + "from expected sheet" + ExpSheetName);
					} else {
						List<Integer> ActresultSet = new ArrayList<Integer>();
						List<Integer> ExpresultSet = new ArrayList<Integer>();
						Actarray.add(ActresultSet);
						Exparray.add(ExpresultSet);
						ActresultSet.add(j);
						ActresultSet.add(i);
						ExpresultSet.add(j);
						ExpresultSet.add(i + 2);
						// ActResults.setCellColor(ActSheetName, j, i, "FAIL");
						// ExpResults.setCellColor(ExpSheetName, j, i+2, "FAIL");
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
					if (splitDataSet.length == 10 && (splitDataSet[0].length() == 1 || splitDataSet[0].length() == 2)
							&& (splitDataSet[1].equals("0") || splitDataSet[1].length() > 4)) {
						// String data1 = splitDataSet[j];

						setCellList_intColumn.add(7 + 0);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[0]);
						setCellList_intColumn.add(7 + 1);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[1]);
						setCellList_intColumn.add(7 + 2);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[2]);
						setCellList_intColumn.add(7 + 3);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[3]);
						setCellList_intColumn.add(7 + 4);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[4]);
						setCellList_intColumn.add(7 + 5);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[5]);
						setCellList_intColumn.add(7 + 6);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add("0");
						setCellList_intColumn.add(7 + 7);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[6]);
						setCellList_intColumn.add(7 + 8);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add("0");
						setCellList_intColumn.add(7 + 9);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[7]);
						setCellList_intColumn.add(7 + 10);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[8]);
						setCellList_intColumn.add(7 + 11);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[9]);
						setCellList_intColumn.add(7 + 12);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[10]);
						

						/*
						 * PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						 * PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						 * PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						 * PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						 * PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						 * PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						 * PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						 * PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, "0");
						 * PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber, splitDataSet[7]);
						 * PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber, "0");
						 * PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[8]);
						 * PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[9]);
						 * PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[10]);
						 * PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[11]);
						 */
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
					if (splitDataSet.length == 10 && (splitDataSet[0].length() == 1 || splitDataSet[0].length() == 2)
							&& (splitDataSet[1].equals("0") || splitDataSet[1].length() > 4)) {
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)

						setCellList_intColumn.add(7 + 0);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[0]);
						setCellList_intColumn.add(7 + 1);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[1]);
						setCellList_intColumn.add(7 + 2);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[2]);
						setCellList_intColumn.add(7 + 3);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[3]);
						setCellList_intColumn.add(7 + 4);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[4]);
						setCellList_intColumn.add(7 + 5);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[5]);
						setCellList_intColumn.add(7 + 6);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add("0");
						setCellList_intColumn.add(7 + 7);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add("0");
						setCellList_intColumn.add(7 + 8);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[6]);
						setCellList_intColumn.add(7 + 9);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[7]);
						setCellList_intColumn.add(7 + 10);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[8]);
						setCellList_intColumn.add(7 + 11);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[9]);
						setCellList_intColumn.add(7 + 12);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[10]);
						

						/*
						 * PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						 * PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						 * PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						 * PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						 * PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						 * PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						 * PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						 * PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, "0");
						 * PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber, "0");
						 * PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber, splitDataSet[7]);
						 * PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[8]);
						 * PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[9]);
						 * PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[10]);
						 * PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[11]);
						 */
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
					if ((splitDataSet.length == 10) && (splitDataSet[0].length() == 1 || splitDataSet[0].length() == 2)
							&& (splitDataSet[1].equals("0") || splitDataSet[1].length() > 4)) {
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)
						setCellList_intColumn.add(7 + 0);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[0]);
						setCellList_intColumn.add(7 + 1);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[1]);
						setCellList_intColumn.add(7 + 2);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[2]);
						setCellList_intColumn.add(7 + 3);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[3]);
						setCellList_intColumn.add(7 + 4);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[4]);
						setCellList_intColumn.add(7 + 5);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[5]);
						setCellList_intColumn.add(7 + 6);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[6]);
						setCellList_intColumn.add(7 + 7);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add("0");
						setCellList_intColumn.add(7 + 8);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add("0");
						setCellList_intColumn.add(7 + 9);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[7]);
						setCellList_intColumn.add(7 + 10);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[8]);
						setCellList_intColumn.add(7 + 11);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[9]);
						/*
						 * setCellList_intColumn.add( 6 + 12); setCellList_intRow.add(rowNumber);
						 * setCellList_Str.add(splitDataSet[10]); setCellList_intColumn.add( 6 + 13);
						 * setCellList_intRow.add(rowNumber); setCellList_Str.add(splitDataSet[11]);
						 */

						/*
						 * PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						 * PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						 * PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						 * PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						 * PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						 * PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						 * PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						 * PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, splitDataSet[7]);
						 * PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber, "0");
						 * PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber, "0");
						 * PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[8]);
						 * PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[9]);
						 * PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[10]);
						 * PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[11]);
						 */
						rowNumber++;
					}
					sb.append(line);
					sb.append(System.lineSeparator());

				}
				line = br.readLine();
				lineNumber++;
				if (line.contains("CURRENT MVA INDEX")) {
					break;
				}
			}
			// lineNumber++;
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
					if (splitDataSet.length == 11 && (splitDataSet[0].length() == 1 || splitDataSet[0].length() == 2)
							&& (splitDataSet[1].equals("0") || splitDataSet[1].length() > 4)) {
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)

						setCellList_intColumn.add(7 + 0);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[0]);
						setCellList_intColumn.add(7 + 1);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[1]);
						setCellList_intColumn.add(7 + 2);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[2]);
						setCellList_intColumn.add(7 + 3);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[3]);
						setCellList_intColumn.add(7 + 4);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[4]);
						setCellList_intColumn.add(7 + 5);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[5]);
						setCellList_intColumn.add(7 + 6);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[6]);
						setCellList_intColumn.add(7 + 7);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add("0");
						setCellList_intColumn.add(7 + 8);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[7]);
						setCellList_intColumn.add(7 + 9);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[8]);
						setCellList_intColumn.add(7 + 10);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[9]);
						setCellList_intColumn.add(7 + 11);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[10]);
						setCellList_intColumn.add(7 + 12);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[11]);
						/*
						 * setCellList_intColumn.add( 6 + 12); setCellList_intRow.add(rowNumber);
						 * setCellList_Str.add(splitDataSet[11]); setCellList_intColumn.add( 6 + 13);
						 * setCellList_intRow.add(rowNumber); setCellList_Str.add(splitDataSet[12]);
						 */

						/*
						 * PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						 * PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						 * PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						 * PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						 * PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						 * PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						 * PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						 * PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, "0");
						 * PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber, splitDataSet[7]);
						 * PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber, splitDataSet[8]);
						 * PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[9]);
						 * PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[10]);
						 * PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[11]);
						 * PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[12]);
						 */
						rowNumber++;
					}
					sb.append(line);
					sb.append(System.lineSeparator());

				}
				line = br.readLine();
				lineNumber++;
				if (line.contains("CURRENT MVA INDEX")) {
					break;
				}
			}
			// lineNumber++;
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
					if (splitDataSet.length == 11 && (splitDataSet[0].length() == 1 || splitDataSet[0].length() == 2)
							&& (splitDataSet[1].equals("0") || splitDataSet[1].length() > 4)) {
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)
						setCellList_intColumn.add(7 + 0);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[0]);
						setCellList_intColumn.add(7 + 1);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[1]);
						setCellList_intColumn.add(7 + 2);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[2]);
						setCellList_intColumn.add(7 + 3);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[3]);
						setCellList_intColumn.add(7 + 4);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[4]);
						setCellList_intColumn.add(7 + 5);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[5]);
						setCellList_intColumn.add(7 + 6);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[6]);
						setCellList_intColumn.add(7 + 7);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[7]);
						setCellList_intColumn.add(7 + 8);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[8]);
						setCellList_intColumn.add(7 + 9);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add("0");
						setCellList_intColumn.add(7 + 10);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[9]);
						setCellList_intColumn.add(7 + 11);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[10]);
						setCellList_intColumn.add(7 + 12);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[11]);
						

						/*
						 * PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						 * PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						 * PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						 * PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						 * PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						 * PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						 * PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						 * PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, splitDataSet[7]);
						 * PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber, splitDataSet[8]);
						 * PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber, "0");
						 * PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[9]);
						 * PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[10]);
						 * PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[11]);
						 * PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[12]);
						 */
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
					if (splitDataSet.length == 11 && (splitDataSet[0].length() == 1 || splitDataSet[0].length() == 2)
							&& (splitDataSet[1].equals("0") || splitDataSet[1].length() > 4)) {
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)
						setCellList_intColumn.add(6 + 0);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[0]);
						setCellList_intColumn.add(6 + 1);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[1]);
						setCellList_intColumn.add(6 + 2);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[2]);
						setCellList_intColumn.add(6 + 3);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[3]);
						setCellList_intColumn.add(6 + 4);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[4]);
						setCellList_intColumn.add(6 + 5);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[5]);
						setCellList_intColumn.add(6 + 6);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[6]);
						setCellList_intColumn.add(6 + 7);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[7]);
						setCellList_intColumn.add(6 + 8);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add("0");
						setCellList_intColumn.add(6 + 9);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[8]);
						setCellList_intColumn.add(6 + 10);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[9]);
						setCellList_intColumn.add(6 + 11);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[10]);
						setCellList_intColumn.add(6 + 12);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[11]);
						setCellList_intColumn.add(6 + 13);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[12]);
						setCellList_intColumn.add(6);
						setCellList_intRow.add(rowNumber);
						setCellList_Str.add(splitDataSet[12]);
						// PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						// PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						// PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						// PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						// PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						// PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						// PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						// PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, splitDataSet[7]);
						// PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber, "0");
						// PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber, splitDataSet[8]);
						// PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[9]);
						// PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[10]);
						// PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[11]);
						// PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[12]);
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

	// When none of these are zero
	public static String ConvertToExcel(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

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
					// System.out.println("Line==" + line);
					String[] splitDataSet = line.split("\\s+");
					// System.out.println("splitData Length=" + splitDataSet.length);
					for (int j = 0; j < splitDataSet.length; j++) {
						if (splitDataSet.length == 12
								&& (splitDataSet[0].length() == 1 || splitDataSet[0].length() == 2)
								&& (splitDataSet[1].equals("0") || splitDataSet[1].length() > 4)) {
							String data1 = splitDataSet[j];
							// System.out.println(data1);
							// setCellData(String sheetName,int colName,int rowNum, String data)
							setCellList_intColumn.add(j + 7);
							setCellList_intRow.add(rowNumber);
							setCellList_Str.add(data1);
							/*
							 * setCellList_intColumn.add(6); setCellList_intRow.add(rowNumber);
							 * setCellList_Str.add("0");
							 */
							// PDFResults.setCellData(ActSheetName, j + 6, rowNumber, data1);
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

	public static void PutPremium(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
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
				if (line.contains("Premium Payment")) {
					String splitValue = line.replaceAll("Premium Payment", " ");
					splitValue = splitValue.replace("$", "");
					setCellList_intColumn.add(6);
					setCellList_intRow.add(rowNumber);
					setCellList_Str.add(splitValue.trim());

				}
				sb.append(line);
				sb.append(System.lineSeparator());

				line = br.readLine();
				lineNumber++;
			}

			String everything = sb.toString();
			// System.out.println(everything);

		} catch (Exception e) {
			e.printStackTrace();

		} finally {
			PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
			System.out.println("text to excel is Done");
			br.close();
		}
	}
}
