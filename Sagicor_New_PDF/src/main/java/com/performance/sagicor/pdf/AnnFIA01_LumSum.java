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

public class AnnFIA01_LumSum {

	/*static String ExpResultsFile = "D:\\Sagicor\\mFolder\\SLIC_SEC_Expected_Results_Single_Page_v29_2018-05-25.xlsx";
	static String ActResultsFile = "D:\\Sagicor\\mFolder\\NewActualresultSagicor1.xlsx";
	static String ExpSheetName = "SEC001";
	static String TextFilepath= "D:\\Sagicor\\mFolder\\PDFToText_SEC001.txt";
	static String ActSheetName = "SEC001";
	static String pdfFilePath= "D:\\Sagicor\\mFolder\\SEC001.pdf";

	public static void main(String[] args) throws Exception {
		ExtentReports extent = new ExtentReports("D:\\mFolder\\test123.html");

		ExtentTest testInst = extent.startTest("test with testcomplte");

		AnnFIA01_LumSumValidation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
				pdfFilePath);
		extent.endTest(testInst);
		extent.flush();
	}*/

	public static void AnnFIA01_LumSumValidation(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {
		//pdftoText(pdfFilePath, TextFilepath);
		Output_HiLowSPFIA14_ReadExcel(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath,
				ActSheetName, pdfFilePath, "LUMP SUM OPTION GUARANTEED VALUES", "MONTHLY INCOME OPTIONS GUARANTEED");
		CompareExcels_AnnFIA01_LumSumValidation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
				TextFilepath, ActSheetName, pdfFilePath);

		//RecordFailResults(testInst,  ExpResultsFile, ActResultsFile,results);
	}

	public static String Output_HiLowSPFIA14_ReadExcel(ExtentTest testInst, String ExpResultsFile,
			String ActResultsFile, String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath,
			String FindValue, String TerminateValue) throws Exception {
		ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
		ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
		ArrayList<String> setCellList_Str = new ArrayList<String>();
		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
		try {
			StringBuilder sb = new StringBuilder();
			// String line1 = br.readLine();
			String line = br.readLine();
			int lineNumber = 1;
			int rowNumber = 3;
			outerloop: for (int i = 0; i <= lineNumber; i++) {
				if (line.contains(FindValue)) {
					// System.out.println("lineNumber==" + lineNumber);
					for (int k = 0; k <= lineNumber; k++) {
						if (line.contains(TerminateValue)) {
							break;
						}
						if (Character.isDigit(line.charAt(0))) {

							String[] splitDataSet = line.split("\\s+");
							// System.out.println("splitData Length=" + splitDataSet.length);

							for (int j = 0; j < splitDataSet.length; j++) {
								if (splitDataSet[0].length() == 2

										&& splitDataSet[1].length() == 2) {
									if (splitDataSet.length == 7) {
										setCellList_intColumn.add(j+53);
										setCellList_intRow.add(rowNumber);
										setCellList_Str.add(splitDataSet[j]);
										//PDFResults.setCellData(ActSheetName, j + 53, rowNumber, splitDataSet[j]);
									}
									if (splitDataSet.length == 6) {

										if (j == 2) {
											setCellList_intColumn.add(j+53);
											setCellList_intRow.add(rowNumber);
											setCellList_Str.add("0");
											setCellList_intColumn.add(j+53+1);
											setCellList_intRow.add(rowNumber);
											setCellList_Str.add(splitDataSet[j]);
											//PDFResults.setCellData(ActSheetName, j + 53, rowNumber, "0");
											/*PDFResults.setCellData(ActSheetName, j + 1 + 53, rowNumber,
													splitDataSet[j]);*/
										} else if (j > 2) {
											setCellList_intColumn.add(j+53+1);
											setCellList_intRow.add(rowNumber);
											setCellList_Str.add(splitDataSet[j]);
											/*PDFResults.setCellData(ActSheetName, j + 1 + 53, rowNumber,
													splitDataSet[j]);*/
										} else if (j < 2) {
											setCellList_intColumn.add(j+53);
											setCellList_intRow.add(rowNumber);
											setCellList_Str.add(splitDataSet[j]);
											//PDFResults.setCellData(ActSheetName, j + 53, rowNumber, splitDataSet[j]);
										}
									}

									if (splitDataSet.length == j + 1) {
										rowNumber++;
									}

									if (line.contains(TerminateValue)) {
										break;
									}
								}
							}

							sb.append(line);
							sb.append(System.lineSeparator());
						}
						line = br.readLine();
						lineNumber++;
					}
					if (line.contains(TerminateValue)) {
						break;
					}
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
			System.out.println("Excel reading end");
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

	public static String CompareExcels_AnnFIA01_LumSumValidation(ExtentTest testInst, String ExpResultsFile,
			String ActResultsFile, String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) {
		List<List<Integer>> Actarray = new ArrayList<List<Integer>>();
		List<List<Integer>> Exparray = new ArrayList<List<Integer>>();
		Xlsx_Reader ExpResults = new Xlsx_Reader(ExpResultsFile);
		Xlsx_Reader ActResults = new Xlsx_Reader(ActResultsFile);
		try {
			for (int i = 2; i <= ActResults.getRowCount(ActSheetName); i++) {
				for (int j = 1; j <= 7; j++) {
					String Actdata = ActResults.getCellFormulaData(ActSheetName, j + 52, i + 1);
					String Expdata = ExpResults.getCellFormulaData(ExpSheetName, j + 79, i + 1);

					if (Actdata.equals(Expdata)) {
						testInst.log(LogStatus.PASS, "Actual value " + Actdata + " from sheet " + ActSheetName
								+ "is matching with " + Expdata + "from expected sheet" + ExpSheetName);
					} else {
						List<Integer> ActresultSet = new ArrayList<Integer>();
						List<Integer> ExpresultSet = new ArrayList<Integer>();
						Actarray.add(ActresultSet);
						Exparray.add(ExpresultSet);
						ActresultSet.add(j+52);
						ActresultSet.add(i+1);
						ExpresultSet.add(j+79);
						ExpresultSet.add(i+1);
						//ActResults.setCellColor(ActSheetName, j+52, i+1, "FAIL");
						//ExpResults.setCellColor(ExpSheetName, j+79, i+1, "FAIL");
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
