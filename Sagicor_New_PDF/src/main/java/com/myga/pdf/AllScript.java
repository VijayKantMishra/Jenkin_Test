package com.myga.pdf;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;

public class AllScript {
	ExtentReports extent = new ExtentReports("D:\\Sagicor\\MYGA\\Extent\\MygaReport_10-07.html");

	@Test(dataProvider = "dp")

	public void f(String TC_Name, String ExpResultsFile, String ActResultsFile, String ExpSheetName,
			String TextFilePath, String ActSheetName, String pdfFilePath) throws Exception {


		ExtentTest testInst = extent.startTest(TC_Name);

		RunAll_MYGA.RunAll(testInst, ExpResultsFile, ActResultsFile, ExpSheetName, TextFilePath, ActSheetName, pdfFilePath);
		

		extent.endTest(testInst);
		extent.flush();
	}

	@DataProvider
	public Object[][] dp() throws Exception {

		return DDT.DDTReader("D:\\Sagicor\\MYGA\\MygaValidation.csv");
	}
}
