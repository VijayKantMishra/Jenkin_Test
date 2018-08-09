package com.performance.sagicor.pdf;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.myga.pdf.DDT;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;

public class SagicorTestRunAll {
	
	ExtentReports extent = new ExtentReports("D:\\Sagicor_New_Final\\SageSecure07-27.html");
	
  @Test(dataProvider = "dp")
  
  public void f(String TC_Name, String ExpResultsFile, String ActResultsFile, String ExpSheetName,
			String TextFilePath, String ActSheetName, String pdfFilePath) throws Exception {


		ExtentTest testInst = extent.startTest(TC_Name);

		RunAll_Sagicor.RunAllFiles(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilePath, ActSheetName, pdfFilePath);
		
		extent.endTest(testInst);
		extent.flush();
	}

	@DataProvider
	public Object[][] dp() throws Exception {

		return DDT.DDTReader("D:\\Sagicor_New_Final\\SagicorPDF_New_Validation.csv");
	}
}
