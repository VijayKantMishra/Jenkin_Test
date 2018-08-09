package com.performance.sagicor.pdf;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class BTFLow14 {
	/*static String ExpResultsFile="D:\\mFolder\\SLIC_SEC_Expected_Results_Single_Page_v29_2018-05-25.xlsx";
	static String ActResultsFile ="D:\\mFolder\\SLIC_FSE62_SEC_Actual_Results_v28_2018-05-03-Demo.xlsx";
	static String ExpSheetName="SEC002";
	static String TextFilepath="D:\\mFolder\\Output_HiLowSPFIA14.txt";
	static String ActSheetName="Output_HiLowSPFIA14_ActResult";
	static String pdfFilePath ="D:\\mFolder\\SEC002.pdf";
	
	
	
	public static void main(String[] args) throws Exception {
		ExtentReports extent = new ExtentReports("D:\\mFolder\\test.html");
		
		ExtentTest testInst = extent.startTest("test with testcomplte");
		
		Output_STF(testInst,ExpResultsFile,ActResultsFile,ExpSheetName,TextFilepath,ActSheetName,pdfFilePath);
		extent.endTest(testInst);
		extent.flush();
	}*/
	public static void Output_STF(ExtentTest testInst,String ExpResultsFile,String ActResultsFile,String ExpSheetName,String TextFilepath,String ActSheetName,String pdfFilePath) throws Exception {
		pdftoText(pdfFilePath, TextFilepath);
		Output_HiLowSPFIA14_ReadExcel(testInst,  ExpResultsFile,
				 ActResultsFile,  ExpSheetName,  TextFilepath,  ActSheetName,  pdfFilePath);
		CompareExcelsStrategy(testInst,  ExpResultsFile,
				 ActResultsFile,  ExpSheetName,  TextFilepath,  ActSheetName,  pdfFilePath);
	}

	public static String Output_HiLowSPFIA14_ReadExcel(ExtentTest testInst, String ExpResultsFile,
			String ActResultsFile, String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

		Xlsx_Reader PDFResults = new Xlsx_Reader(
				ActResultsFile);
		BufferedReader br = new BufferedReader(
				new FileReader(TextFilepath));
		try {
			StringBuilder sb = new StringBuilder();
			// String line1 = br.readLine();
			String line = br.readLine();
			int lineNumber = 1;
			int rowNumber = 2;

			for (int i = 0; i <= lineNumber; i++) {
				// System.out.println("lineNumber==" + lineNumber);
				if (Character.isDigit(line.charAt(0))) {
					System.out.println("Line==" + line);
					String[] splitDataSet = line.split("\\s+");
					// System.out.println("splitData Length=" + splitDataSet.length);

					for (int j = 0; j < splitDataSet.length; j++) {
						if (splitDataSet[0].length() == 1
								|| splitDataSet[0].length() == 2 && (splitDataSet[1].length() == 10)) {
							String data1 = splitDataSet[j];
							System.out.println(data1);
							// setCellData(String sheetName,int colName,int rowNum, String data)
							PDFResults.setCellData(ActSheetName, j, rowNumber, data1);
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

	public static String CompareExcelsStrategy(ExtentTest testInst, String ExpResultsFile,
			String ActResultsFile, String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) {
		try {
			Xlsx_Reader ExpResults = new Xlsx_Reader(
					ExpResultsFile);
			Xlsx_Reader ActResults = new Xlsx_Reader(
					ActResultsFile);

			// getCellData(String sheetName,int colNum,int rowNum)
			//System.out.println("ActResults.getColumnCount(\"Sheet1\")" + ActResults.getColumnCount("Sheet1"));
			//System.out.println("ActResults.getRowCount(\"Sheet1\")" + ActResults.getRowCount("Sheet1"));
			for (int i = 2; i <= ActResults.getRowCount(ActSheetName); i++) {
				for (int j = 1; j < ActResults.getColumnCount(ActSheetName); j++) {
					String Actdata = ActResults.getCellFormulaData(ActSheetName, j, i);
					String Expdata = ExpResults.getCellFormulaData(ExpSheetName, j+42, i + 2);
					System.out.println("ActData*************" + Actdata);
					System.out.println("ExpData*************" + Expdata);
					if (Actdata.equals(Expdata)) {
						testInst.log(LogStatus.PASS, Actdata, "validation is pass");
					} else {
						ActResults.setCellColor(ActSheetName, j, i, "FAIL");
						ExpResults.setCellColor(ExpSheetName, j+42, i+2, "FAIL");
						testInst.log(LogStatus.FAIL,
								"Validation is failed at: column " + j + " at row: " + i + " for sheet name: "
										+ ExpSheetName + "Actual result is : " + Actdata + "Expected result is : "
										+ Expdata);
					}
					// return "PASS";
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		}
		System.out.println("File compare is Done");
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
