package com.performance.sagicor.pdf;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;

public class RunAll_Sagicor {

/*	static String ExpResultsFile = "D:\\Sagicor_New_Final\\SLIC_SEC_Expected_Results_Single_Page_v29_2018-07-19 - Copy.xlsx";
	static String ActResultsFile = "D:\\Sagicor_New_Final\\NewActualresultSagicor_New.xlsx";
	static String ExpSheetName = "SEC014";
	static String TextFilepath= "D:\\Sagicor_New_Final\\PDFToText_SEC001.txt";
	static String ActSheetName = "SEC001";
	static String pdfFilePath= "D:\\Sagicor_New_Final\\SEC014.pdf";
	static String FindValue;
	static String TerminateValue;
	
	public static void main(String[] args) throws Exception {
		ExtentReports extent = new ExtentReports("D:\\mFolder\\resultNewFile_2Files_New.html");

		ExtentTest testInst = extent.startTest("test with testcomplte");

		RunAllFiles(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName, pdfFilePath);
		extent.endTest(testInst);
		extent.flush(); 
	}*/

	public static void RunAllFiles(ExtentTest testInst,String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception {
	
	
	StrategyOutPut_NewActual.Page4StrategyOutPutValidation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
				TextFilepath, ActSheetName, pdfFilePath);
		Output_30_FIA01_NewActual.output_30_FIAValidation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
				TextFilepath, ActSheetName, pdfFilePath);
		MVADemoFIA14_NewActual.MVADemoFIA14Validation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
				TextFilepath, ActSheetName, pdfFilePath);
	HiLowBTF_NewActual.Output_HiLowBTFIA14Validation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
				TextFilepath, ActSheetName, pdfFilePath);
		HiLowSPF14_NewActual.Output_HiLowSPFIA14Validation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
				TextFilepath, ActSheetName, pdfFilePath);
		Output_SummaryFIA01.Output_SummaryFIA01validation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
				TextFilepath, ActSheetName, pdfFilePath);
		/*AnnFIA01_LumSum.AnnFIA01_LumSumValidation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath,
				ActSheetName, pdfFilePath);*/
		AnnFIA01_MonthlyIncome.AnnFIA01_LumSumValidation(testInst, ExpResultsFile, ActResultsFile, ExpSheetName,
			TextFilepath, ActSheetName, pdfFilePath);
	}
	
	
}
