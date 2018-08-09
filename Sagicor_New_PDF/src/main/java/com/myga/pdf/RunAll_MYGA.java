package com.myga.pdf;

import java.io.IOException;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;

public class RunAll_MYGA {

	
	static String ExpResultsFile = "D:\\Sagicor\\MYGA\\SLIC_MYGA_Expected_Results_Single_Page_v28_2018-06-26.xlsx";
	static String ActResultsFile = "D:\\Sagicor\\MYGA\\Myga_ActualResult5.xlsx";
	static String ExpSheetName = "myga008";
	static String TextFilepath = "D:\\Sagicor\\MYGA\\PDFToText_myga001.txt";
	static String ActSheetName = "myga008";
	static String pdfFilePath = "D:\\Sagicor\\MYGA\\myga008.pdf";

	
	public static void main(String[] args) throws Exception {
		ExtentReports extent = new ExtentReports("D:\\mFolder\\test.html");

		ExtentTest testInst = extent.startTest("test with testcomplte");

		RunAll(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilepath, ActSheetName,
				pdfFilePath);
		extent.endTest(testInst);
		extent.flush();
	}
	
	public static void RunAll(ExtentTest testInst, String ExpResultsFile, String ActResultsFile,
			String ExpSheetName, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception, Exception{
		
		DEMONSTRATION_HYPOTHETICAL_MVA.output_30_FIAValidation( testInst,  ExpResultsFile,  ActResultsFile,
				 ExpSheetName,  TextFilepath,  ActSheetName,  pdfFilePath);
		SETTLEMENT_OPTIONS_MonthlyIncome.AnnFIA01_LumSumValidation( testInst,  ExpResultsFile,  ActResultsFile,
				 ExpSheetName,  TextFilepath,  ActSheetName,  pdfFilePath);
		HYPOTHETICAL_ILLUSTRATED_VALUES.output_30_FIAValidation(testInst,  ExpResultsFile,  ActResultsFile,
				 ExpSheetName,  TextFilepath,  ActSheetName,  pdfFilePath);
		SETTLEMENT_OPTIONS_LumSum.AnnFIA01_LumSumValidation(testInst,  ExpResultsFile,  ActResultsFile,
				 ExpSheetName,  TextFilepath,  ActSheetName,  pdfFilePath);
				
			} 
}
