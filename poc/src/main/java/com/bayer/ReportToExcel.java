package com.bayer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;


public class ReportToExcel {
	
	
	private final static String File_Name_In = "E:\\sitelistinput.xlsx";
	private final static String File_Name_Out = "E:\\sitelistout.xlsx";

//	private final static String File_Name_In = "C:\\Users\\EPDZX\\OneDrive - Bayer\\Desktop\\selenium\\sitelistinput.xlsx";
//	private final static String File_Name_Out = "C:\\Users\\EPDZX\\OneDrive - Bayer\\Desktop\\\\selenium\\siteout.xlsx";
	static WebDriver driver;

	public String test() throws IOException {

//Chrome Driver
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\Abhishek\\Desktop\\selenium\\instal\\chromedriver_win32\\chromedriver.exe");
		driver = new ChromeDriver();
//Firefox Driver
//		System.setProperty("webdriver.gecko.driver",
//				"C:\\Users\\EPDZX\\OneDrive - Bayer\\Desktop\\selenium\\geckodriver-v0.24.0-win64\\geckodriver.exe");
//		driver = new FirefoxDriver();
//IE Driver
//		System.setProperty("webdriver.ie.driver",
//			"C:\\Users\\EPDZX\\OneDrive - Bayer\\Desktop\\selenium\\IEDriverServer_x64_3.14.0\\IEDriverServer.exe");
//		driver = new InternetExplorerDriver();

		FileInputStream excelFileInput = new FileInputStream(new File(File_Name_In));
		Workbook workbookIn = new XSSFWorkbook(excelFileInput);
		org.apache.poi.ss.usermodel.Sheet datatypeSheet = workbookIn.getSheetAt(0);
		workbookIn.close();
		Iterator<Row> iterator = datatypeSheet.iterator();
		int sitelistsize = datatypeSheet.getLastRowNum();
		XSSFWorkbook workbook = new XSSFWorkbook();
		// Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Content Check");
		
		

		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] { "SiteName", "Privacy Statement", "Imprint", "Contact Us", "Conditions of Use",
				"Bayer Group/Bayer Global", "LMR Number","Bayer Logo" });
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();

			while (cellIterator.hasNext()) {

				Cell currentCell = cellIterator.next();
				String sitename = currentCell.getStringCellValue();

				String isPrivacyPresnt = "";
				String isImprintUsed = "";
				String isContactUsUsed = "";
				String isConditionsofUseUsed = "";
				String isBayerGroupUsed = "";
				String isLMRNumberUsed = "";
				String isBayerlogoUsed = "";

				driver.get(sitename);
				Dimension n = new Dimension(360, 592);
				driver.manage().window().setSize(n);
				// Privacy Policy,Imprint, Bayer Logo,LMR Number,Bayer Group/ Bayer
				// Global,Contact Us,General Conditions of Use


//Privacy Statement/Privacy Policy Check 

				String[] strArrayPrivacystatement = new String[] { "Privacy Policy", "Privacy Statement","PRIVACY STATEMENT" };
				int lengthofPrivacystatement = strArrayPrivacystatement.length;
				for (int j = 0; j < lengthofPrivacystatement; j++) {
					try {
						driver.findElement(By.linkText(strArrayPrivacystatement[j])).isDisplayed();
						isPrivacyPresnt = "YES";
						if (isPrivacyPresnt == "YES") {
							break;
						}
					} catch (NoSuchElementException e) {
						isPrivacyPresnt = "NO";
					}
				}
//Imprint Check 

				String[] strArrayImprint = new String[] { "Imprint", "IMPRINT" };
				int lengthofImprint = strArrayImprint.length;
				for (int j = 0; j < lengthofImprint; j++) {
					try {
						driver.findElement(By.linkText(strArrayImprint[j])).isDisplayed();
						isImprintUsed = "YES";
						if (isImprintUsed == "YES") {
							break;
						}
					} catch (NoSuchElementException e) {
						isImprintUsed = "NO";
					}
				}

// Contact Us Check 				
				String[] strArraycontactus = new String[] { "Contact Us", "CONTACT US", "Contact us" };
				int lengthofcontactus = strArraycontactus.length;
				for (int j = 0; j < lengthofcontactus; j++) {
					try {
						driver.findElement(By.partialLinkText(strArraycontactus[j])).isDisplayed();
						isContactUsUsed = "YES";
						if (isContactUsUsed == "YES") {
							break;
						}
					} catch (NoSuchElementException e) {
						isContactUsUsed = "NO";
					}
				}
// Conditions of Use Check 				
				String[] strArrayconditionosuse = new String[] { "Conditions of Use", "CONDITIONS OF USE" };
				int lengthofconditionosuse = strArrayconditionosuse.length;
				for (int j = 0; j < lengthofconditionosuse; j++) {
					try {
						driver.findElement(By.linkText(strArrayconditionosuse[j])).isDisplayed();
						isConditionsofUseUsed = "YES";
						if (isConditionsofUseUsed == "YES") {
							break;
						}
					} catch (NoSuchElementException e) {
						isConditionsofUseUsed = "NO";
					}
				}
// Bayer Group/Bayer Global Check 				
				String[] strArraybayerglgobal = new String[] { "Bayer Group", "BAYER GLOBAL","Bayer Global" };
				int lengthofbayerglgobal = strArraybayerglgobal.length;
				for (int j = 0; j < lengthofbayerglgobal; j++) {
					try {
						//driver.findElement(By.linkText(strArraybayerglgobal[j])).isDisplayed();
						driver.findElement(By.partialLinkText(strArraybayerglgobal[j])).isDisplayed();
						isBayerGroupUsed = "YES";
						if (isBayerGroupUsed == "YES") {
							break;
						}
					}

					catch (NoSuchElementException e) {
						isBayerGroupUsed = "NO";
					}
				}
//LMR Value
				try {
					boolean lmrvalue;
					lmrvalue = driver.getPageSource().contains("L.CA.MKT.CC");
					if (lmrvalue == true) {
						isLMRNumberUsed = "YES";
					} else {
						isLMRNumberUsed = "NO";
					}
				} catch (NoSuchElementException e) {
				}
//Bayer Logo			
				try 
				{
					String[] strArraybayerlogo = new String[] { "http://www.bayer.com", "https://www.bayer.com","http://www.bayer.com/","/en/homepage.aspx" };
					int lengthofbayerlogo = strArraybayerlogo.length;
					boolean link=true;
					for (int j = 0; j < lengthofbayerlogo; j++) 
					{
						List<WebElement> list = driver.findElements(By.xpath("//*[@href='"+strArraybayerlogo[j]+"']//img"));
					
					for (WebElement element : list)
					{
					   link = element.getAttribute("src").isEmpty();
					   if(link==false)
					    {
					    isBayerlogoUsed="YES";
					    }
					 }
					if(link==true)
				    {
						isBayerlogoUsed = "NO";
				    }
					
				}
				 
				
				}
				catch (NoSuchElementException e){}
				
				data.put(sitename, new Object[] { sitename, isPrivacyPresnt, isImprintUsed, isContactUsUsed,
						isConditionsofUseUsed, isBayerGroupUsed, isLMRNumberUsed,isBayerlogoUsed });
			}

		}

		// Iterate over data and write to sheet
		Set<String> keyset = data.keySet();
		
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);


				// Setting style only for header
				if (rownum == 1) {
					// CellStyle style=null;
					// Creating a font
					
					XSSFFont font = workbook.createFont();
					font.setFontHeightInPoints((short) 10);
					font.setFontName("Arial");
					font.setColor(IndexedColors.WHITE.getIndex());

					font.setBold(true);
					font.setItalic(false);
					CellStyle style = workbook.createCellStyle();
					// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
					// style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//
					style.setAlignment(HorizontalAlignment.CENTER);
					style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

					// Setting font to style
					style.setFont(font);
					style.setWrapText(false);
					style.setBorderBottom(BorderStyle.THIN);
					style.setBorderLeft(BorderStyle.THIN);
					style.setBorderRight(BorderStyle.THIN);
					style.setAlignment(HorizontalAlignment.CENTER);

					// Setting cell style
					cell.setCellStyle(style);
					
					sheet.autoSizeColumn(1);sheet.autoSizeColumn(2);sheet.autoSizeColumn(3);sheet.autoSizeColumn(4);
					sheet.autoSizeColumn(5);sheet.autoSizeColumn(6);sheet.autoSizeColumn(7);
					
				} else {
					
					sheet.autoSizeColumn(0);
					XSSFFont font2 = workbook.createFont();
					font2.setFontHeightInPoints((short) 10);
					font2.setFontName("Arial");
					font2.setColor(IndexedColors.BLACK.getIndex());

					font2.setBold(false);
					font2.setItalic(false);

					CellStyle style2 = workbook.createCellStyle();

					// Setting font to style
					style2.setBorderBottom(BorderStyle.THIN);
					style2.setBorderLeft(BorderStyle.THIN);
					style2.setBorderRight(BorderStyle.THIN);
					style2.setAlignment(HorizontalAlignment.CENTER);
					style2.setFont(font2);
					style2.setWrapText(false);
					if(obj=="YES")
					{
						style2.setFillForegroundColor(IndexedColors.BRIGHT_GREEN1.getIndex());
						style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					}
					if(obj=="NO")
					{
						style2.setFillForegroundColor(IndexedColors.RED1.getIndex());
						style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					}
					if(obj!="NO"&&obj!="YES")
					{
						style2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
						style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					}
					

					// Setting cell style
					cell.setCellStyle(style2);
					
				}
			}
		}
		try {
			// Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File(File_Name_Out));
			workbook.write(out);
			out.close();
			System.out.println("Excel file created");
			driver.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
return "Pass";
	}

	public static void main(String[] args) throws IOException {
		ReportToExcel myobj = new ReportToExcel();
		myobj.test();
	}

}