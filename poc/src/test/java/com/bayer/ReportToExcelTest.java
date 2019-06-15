package com.bayer;

import java.io.IOException;

import org.openqa.selenium.support.PageFactory;
import org.testng.Assert;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class ReportToExcelTest extends TestNgTestBase {
	private ReportToExcel reportExcelTest;
	 @BeforeMethod
	  public void initPageObjects() {
		 reportExcelTest = PageFactory.initElements(driver, ReportToExcel.class);
	  }

	  @Test
	  public void testHomePageHasAHeader() {
	    driver.get(baseUrl);
	    try {
			Assert.assertTrue("Pass".equals(reportExcelTest.test()));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	  }

}
