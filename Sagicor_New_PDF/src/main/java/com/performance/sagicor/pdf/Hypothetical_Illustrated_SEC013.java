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

public class Hypothetical_Illustrated_SEC013 {
	
	static String ExpResultsFile = "D:\\mFolder\\SLIC_SEC_Expected_Results_Single_Page_v29_2018-05-25.xlsx";
	static String ActResultsFile = "D:\\mFolder\\NewActualresult.xlsx";
	static String ExpSheetName = "SEC010";
	static String TextFilepath = "D:\\mFolder\\PDFToText_SEC010.txt";
	static String ActSheetName = "SEC010";
	static String pdfFilePath = "D:\\mFolder\\SEC014.pdf";

public static void main(String[] args) throws Exception {
		ExtentReports extent = new ExtentReports("D:\\mFolder\\test1.html");

		ExtentTest testInst = extent.startTest("test with testcomplte");

		Output_SummaryFIA01validation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath,
				ActSheetName, pdfFilePath);
		extent.endTest(testInst);
		extent.flush();
	}
	
	public static void Output_SummaryFIA01validation(ExtentTest testInst, String ExpResultsFile,
			String ActResultsFile, String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception, IOException {
		pdftoText(pdfFilePath,TextFilepath);
		
		ConvertToExcel( testInst,  ExpResultsFile,
				 ActResultsFile,  ExpSheetName,  TextFilepath,  ActSheetName,  pdfFilePath );
		/*CompareExcels( testInst,  ExpResultsFile,
				 ActResultsFile,  ExpSheetName,  TextFilepath,  ActSheetName,  pdfFilePath );*/
	}
	
	public static String CompareExcels(ExtentTest testInst, String ExpResultsFile,
			String ActResultsFile, String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath ) {
		try {
			Xlsx_Reader ExpResults = new Xlsx_Reader(ExpResultsFile);
			Xlsx_Reader ActResults = new Xlsx_Reader(ActResultsFile);

			for (int i = 3; i <= ActResults.getRowCount(ActSheetName); i++) {
				for (int j = 6; j <= 19; j++) {
					String Actdata = ActResults.getCellFormulaData(ActSheetName, j, i);
					String Expdata = ExpResults.getCellFormulaData(ExpSheetName, j, i+2);
					System.out.println("ActData*************" + Actdata);
					System.out.println("ExpData*************" + Expdata);
					if (Actdata.equals(Expdata)) {
						testInst.log(LogStatus.PASS,"Values are matching");
					} else {
						ActResults.setCellColor(ActSheetName, j, i, "FAIL");
						ExpResults.setCellColor(ExpSheetName, j, i+2, "FAIL");
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

	public static String ConvertToExcel(ExtentTest testInst, String ExpResultsFile,
			String ActResultsFile, String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {

		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
		try {
			StringBuilder sb = new StringBuilder();
			// String line1 = br.readLine();
			String line = br.readLine();
			int lineNumber = 1;
			int rowNumber = 3;

			for (int i = 0; i <= lineNumber; i++) {
				//System.out.println("lineNumber==" + lineNumber);
				if (Character.isDigit(line.charAt(0))) {
					//System.out.println("Line==" + line);
					String[] splitDataSet = line.split("\\s+");
					//System.out.println("splitData Length=" + splitDataSet.length);

					for (int j = 0; j < splitDataSet.length; j++) {
						if (splitDataSet[0].length() == 10 || (splitDataSet[2].equals("0") && Integer.parseInt(splitDataSet[2])>1000)){
 {
							String data1 = splitDataSet[j];
							//System.out.println(data1);
							// setCellData(String sheetName,int colName,int rowNum, String data)
							PDFResults.setCellData(ActSheetName, j+6, rowNumber, data1);
							if (splitDataSet.length == j + 1) {
								rowNumber++;
							}
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

	public static String pdftoText(String pdfFilePath,String TextFilepath) throws InterruptedException, IOException {
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
