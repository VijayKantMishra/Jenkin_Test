package com.performance.sagicor.pdf;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class Test_Pdf_001 {
	static String ExpResultsFile = "D:\\mFolder\\SLIC_SEC_Expected_Results_Single_Page_v29_2018-05-25.xlsx";
	static String ActResultsFile = "D:\\mFolder\\NewActualresult.xlsx";
	static String ExpSheetName = "SEC002";
	static String TextFilepath = "D:\\mFolder\\PDFToText_SEC001.txt";
	static String ActSheetName = "SEC002";
	static String pdfFilePath = "D:\\mFolder\\SEC002.pdf";

	static String restLineValue2;
	static String restLineValue1;
	static String restLineValue;
	static String[] splitDataSet;
	static String[] splitDataSet1;
	static String[] splitDataSet2;
	static ArrayList<String> arry = new ArrayList<String>();

	public static void main(String[] args) throws Exception {
		ExtentReports extent = new ExtentReports("D:\\Sagicor\\mFolder\\test.html");

		ExtentTest testInst = extent.startTest("test with testcomplte");

		output_30_FIAValidation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
				pdfFilePath);
		extent.endTest(testInst);
		extent.flush();
	}

	public static void output_30_FIAValidation(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath)
					throws Exception, IOException {
		pdftoText(pdfFilePath, TextFilepath);

		String result = GetStrategy(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
				pdfFilePath, "Declared Rate Strategy", 6);
		String DRS = result.split("&")[0];
		String SNPS = result.split("&")[1];
		String GMIS = result.split("&")[2];

		System.out.println("Result1=" + result.split("&")[0]);
		System.out.println("Result2=" + result.split("&")[1]);
		System.out.println("Result3=" + result.split("&")[2]);
		System.out.println("Value");
		System.out.println(!DRS.startsWith("0%") && !SNPS.equals("0%") && !GMIS.equals("0%"));
		System.out.println(DRS.equals("0%")  && SNPS.equals("100%") && GMIS.equals("0%"));

		if (!DRS.startsWith("0%") && !SNPS.equals("0%") && !GMIS.equals("0%")) {
			result = ConvertToExcel(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);

		} else if (DRS.equals("0%")  && ! SNPS.equals("0%") && GMIS.equals("0%")) {

			result = DecleredRateAndGlobalIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);
		}else if(DRS.equals("0%")  && SNPS.equals("0%") && ! GMIS.equals("0%")) {

			result = DecleredRateAndSNPIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);
		}else if(! DRS.equals("0%")  && SNPS.equals("0%") &&  GMIS.equals("0%")) {

			result = SNPANDGlobalIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);
		}else if( DRS.equals("0%")  && ! SNPS.equals("0%") &&  ! GMIS.equals("0%")) {

			result = DecleredIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);
		}else if(! DRS.equals("0%")  && ! SNPS.equals("0%") &&   GMIS.equals("0%")) {

			result = GlobalIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);
		} else if (! DRS.equals("0%")  &&  SNPS.equals("0%") &&   ! GMIS.equals("0%")) {
			result = SNPIsZero(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
					pdfFilePath);
		}
		// ConvertToExcel(testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
		// TextFilepath, ActSheetName, pdfFilePath,"S&P 500® Index Strategy",7);

		HashMap<Integer, String[]>	results = CompareExcels( testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
				TextFilepath, ActSheetName, pdfFilePath );

		
		RecordFailResults(testInst,results);

	}
	
	public static boolean RecordFailResults(ExtentTest testInst,HashMap<Integer, String[]> results) {
		try {
			Xlsx_Reader ExpResults = new Xlsx_Reader(ExpResultsFile);
			Xlsx_Reader ActResults = new Xlsx_Reader(ActResultsFile);
			
			for (Map.Entry<Integer, String[]> entry : results.entrySet())
			{
				String[] actData = entry.getValue()[0].split("#");
				ActResults.setRedColor(actData[0],Integer.parseInt(actData[1]),Integer.parseInt(actData[2]));

				String[] expData = entry.getValue()[1].split("#");
				ExpResults.setRedColor(expData[0],Integer.parseInt(expData[1]),Integer.parseInt(expData[2]));
				testInst.log(LogStatus.FAIL,
						"Validation is failed at: column  sheet name: "
								+ expData[0] + "Actual result is : " + actData[3] + "Expected result is : "
								+ expData[3]);
			}

			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}
	

	public static HashMap<Integer, String[]> CompareExcels(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) {
		HashMap<Integer, String[]> results = new HashMap<Integer, String[]>();
		try {
			Xlsx_Reader ExpResults = new Xlsx_Reader(ExpResultsFile);
			Xlsx_Reader ActResults = new Xlsx_Reader(ActResultsFile);
			
			int counter = 1;
			for (int i = 3; i <= ActResults.getRowCount(ActSheetName); i++) {
				for (int j = 6; j <= 19; j++) {
					String Actdata = ActResults.getCellFormulaData(ActSheetName, j, i);
					String Expdata = ExpResults.getCellFormulaData(ExpSheetName, j, i + 2);
					System.out.println("ActData*************" + Actdata);
					System.out.println("ExpData*************" + Expdata);
					if (Actdata.equals(Expdata)) {
						testInst.log(LogStatus.PASS, "Values are matching");
					} else {
						String[] failures = new String[2];
						arry.add(i+ "RESULT IS "+ j);
						failures[0]= ActSheetName+"#"+j+"#"+i+"#"+Actdata;
						if(Expdata==null || Expdata.equals("")) {
							failures[1]= ExpSheetName+"#"+j+"#"+(i+2)+"#EMPTY";
						}
						failures[1]= ExpSheetName+"#"+j+"#"+(i+2)+"#"+Expdata;
						results.put(counter, failures);
						counter++;

						//ActResults.setCellColor(ActSheetName, j, i, "FAIL");
						//ExpResults.setCellColor(ExpSheetName, j, i + 2, "FAIL");
						//testInst.log(LogStatus.FAIL,
						//								"Validation is failed at: column " + j + " at row: " + i + " for sheet name: "
						//										+ ExpSheetName + "Actual result is : " + Actdata + "Expected result is : "
						//										+ Expdata);
					}
				}
			}
			return results ;
		} catch (Exception e) {
			System.out.println("File compare is Exception");
			e.printStackTrace();
			return null;
		}
	}	
	

	public static void PrintResult(String ExpResultsFile,String ActResultsFile) {

		Xlsx_Reader ExpResults = new Xlsx_Reader(ExpResultsFile);
		Xlsx_Reader ActResults = new Xlsx_Reader(ActResultsFile);
		for (int i = 0; i <=arry.size(); i++) {
			for (int j = 0; j <= 13; j++) {
				ActResults.setCellColor(ActSheetName, j+6, i+3, "FAIL");
			}
		}

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
					//String data1 = splitDataSet[j];


					//for (int j = 0; j < splitDataSet.length; j++) {
					if (splitDataSet[0].length() == 10
							|| (splitDataSet[2].equals("0") && Integer.parseInt(splitDataSet[2]) > 1000)) {
						//String data1 = splitDataSet[j];
						PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, "0");
						PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber, splitDataSet[7]);
						PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber, "0");
						PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[8]);
						PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[9]);
						PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[10]);
						PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[11]);
						rowNumber++;
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)
						/*if (j == 8 + 1) {
								PDFResults.setCellData(ActSheetName, j + 6, rowNumber, "0");
								PDFResults.setCellData(ActSheetName, j + 6 + 1, rowNumber, splitDataSet[j - 1]);
							} else if (j > 8 + 1) {
								PDFResults.setCellData(ActSheetName, j + 6 + 1, rowNumber, splitDataSet[j - 1]);
								PDFResults.setCellData(ActSheetName, j + 6 + 2, rowNumber, splitDataSet[j]);
							} else if (j == 7 + 1) {

								  PDFResults.setCellData(ActSheetName, j + 6, rowNumber, "0");

								PDFResults.setCellData(ActSheetName, j + 6 + 1, rowNumber, data1);
							} else if (j == 6 + 1) {
								PDFResults.setCellData(ActSheetName, j + 6, rowNumber, "0");
								PDFResults.setCellData(ActSheetName, j + 6 + 1, rowNumber, data1);
							} 

								  else if (j < 6 + 1) {
								PDFResults.setCellData(ActSheetName, j + 6, rowNumber, data1);
							}*/

						//							if (splitDataSet.length == j + 1) {
						//								rowNumber++;
						//							}

					}
					//}
					sb.append(line);
					sb.append(System.lineSeparator());
				}
				line = br.readLine();
				lineNumber++;
			}
			String everything = sb.toString();
			// System.out.println(everything);
			System.out.println("text to excel is Done");
			return "PASS";
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		} finally {
			br.close();
		}
	}
	// When Decleared rate and SnP is zero and Global is not zero
	public static String DecleredRateAndSNPIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
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

					///for (int j = 0; j < splitDataSet.length; j++) {
					if (splitDataSet[0].length() == 10
							|| (splitDataSet[2].equals("0") && Integer.parseInt(splitDataSet[2]) > 1000)) {
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)
						PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, "0");
						PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber, "0");
						PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber, splitDataSet[7]);
						PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[8]);
						PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[9]);
						PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[10]);
						PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[11]);
						rowNumber++;
					}
					//}
					sb.append(line);
					sb.append(System.lineSeparator());
				}
				line = br.readLine();
				lineNumber++;
			}
			String everything = sb.toString();
			// System.out.println(everything);
			System.out.println("text to excel is Done");
			return "PASS";
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		} finally {
			br.close();
		}
	}

	// When Decleared rate is not zero and S&P,Global is zero
	public static String SNPANDGlobalIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
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

					/*for (int j = 0; j < splitDataSet.length; j++) {*/
					if (splitDataSet[0].length() == 10
							|| (splitDataSet[2].equals("0") && Integer.parseInt(splitDataSet[2]) > 1000)) {
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)

						PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, splitDataSet[7]);
						PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber, "0");
						PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber,"0" );
						PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[8]);
						PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[9]);
						PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[10]);
						PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[11]);
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
			System.out.println("text to excel is Done");
			return "PASS";
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		} finally {
			br.close();
		}
	}

	// When Decleared rate is zero and S&P,Global is not zero
	public static String DecleredIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
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

					/*for (int j = 0; j < splitDataSet.length; j++) {*/
					if (splitDataSet[0].length() == 10
							|| (splitDataSet[2].equals("0") && Integer.parseInt(splitDataSet[2]) > 1000)) {
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)

						PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, "0");
						PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber, splitDataSet[7]);
						PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber,splitDataSet[8] );
						PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[9]);
						PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[10]);
						PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[11]);
						PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[12]);
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
			System.out.println("text to excel is Done");
			return "PASS";
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		} finally {
			br.close();
		}
	}
	// When Decleared rate and S&P is not zero,Global is zero 
	public static String GlobalIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
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

					/*for (int j = 0; j < splitDataSet.length; j++) {*/
					if (splitDataSet[0].length() == 10
							|| (splitDataSet[2].equals("0") && Integer.parseInt(splitDataSet[2]) > 1000)) {
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)

						PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, splitDataSet[7]);
						PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber,splitDataSet[8] );
						PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber,"0" );
						PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[9]);
						PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[10]);
						PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[11]);
						PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[12]);
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
			System.out.println("text to excel is Done");
			return "PASS";
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		} finally {
			br.close();
		}
	}


	// When Decleared rate and  Global is not zero and S&P is  zero.
	public static String SNPIsZero(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
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

					/*for (int j = 0; j < splitDataSet.length; j++) {*/
					if (splitDataSet[0].length() == 10
							|| (splitDataSet[2].equals("0") && Integer.parseInt(splitDataSet[2]) > 1000)) {
						// System.out.println(data1);
						// setCellData(String sheetName,int colName,int rowNum, String data)

						PDFResults.setCellData(ActSheetName, 6 + 0, rowNumber, splitDataSet[0]);
						PDFResults.setCellData(ActSheetName, 6 + 1, rowNumber, splitDataSet[1]);
						PDFResults.setCellData(ActSheetName, 6 + 2, rowNumber, splitDataSet[2]);
						PDFResults.setCellData(ActSheetName, 6 + 3, rowNumber, splitDataSet[3]);
						PDFResults.setCellData(ActSheetName, 6 + 4, rowNumber, splitDataSet[4]);
						PDFResults.setCellData(ActSheetName, 6 + 5, rowNumber, splitDataSet[5]);
						PDFResults.setCellData(ActSheetName, 6 + 6, rowNumber, splitDataSet[6]);
						PDFResults.setCellData(ActSheetName, 6 + 7, rowNumber, splitDataSet[7]);
						PDFResults.setCellData(ActSheetName, 6 + 8, rowNumber,"0" );
						PDFResults.setCellData(ActSheetName, 6 + 9, rowNumber,splitDataSet[8] );
						PDFResults.setCellData(ActSheetName, 6 + 10, rowNumber, splitDataSet[9]);
						PDFResults.setCellData(ActSheetName, 6 + 11, rowNumber, splitDataSet[10]);
						PDFResults.setCellData(ActSheetName, 6 + 12, rowNumber, splitDataSet[11]);
						PDFResults.setCellData(ActSheetName, 6 + 13, rowNumber, splitDataSet[12]);
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
			System.out.println("text to excel is Done");
			return "PASS";
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		} finally {
			br.close();
		}
	}



	// When none of these are zero
	public static String ConvertToExcel(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

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
				if (Character.isDigit(line.charAt(0))) {
					// System.out.println("Line==" + line);
					String[] splitDataSet = line.split("\\s+");
					// System.out.println("splitData Length=" + splitDataSet.length);

					for (int j = 0; j < splitDataSet.length; j++) {
						if (splitDataSet[0].length() == 10
								|| (splitDataSet[2].equals("0") && Integer.parseInt(splitDataSet[2]) > 1000)) {
							String data1 = splitDataSet[j];
							// System.out.println(data1);
							// setCellData(String sheetName,int colName,int rowNum, String data)
							PDFResults.setCellData(ActSheetName, j + 6, rowNumber, data1);
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
			System.out.println("text to excel is Done");
			return "PASS";
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		} finally {
			br.close();
		}
	}
}