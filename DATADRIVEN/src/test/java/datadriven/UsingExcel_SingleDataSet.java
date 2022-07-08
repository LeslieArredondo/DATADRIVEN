package datadriven;

import java.util.concurrent.TimeUnit;

import org.apache.xml.security.utils.Constants;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import utilities.ExcelUtility;
import utilities.ExcelUtility_SingleDataSet;

public class UsingExcel_SingleDataSet {
	private WebDriver driver;
	
	@BeforeClass
	public void setUp() throws Exception {
		driver = new ChromeDriver();
		
		//Maximize the browser's window
				driver.manage().window().maximize();
				driver.manage().timeouts().implicitWait(10, TimeUnit.SECONDS);
				driver.get(Constants.URL);
				driver.findElement(By.xpath("//span[text()='Learn Now']")).click();
				//Tell the code about the location of Excel file
				ExcelUtility.setExcelFile(Constants.File_Path + Constants.File_Name, "Sheet1");
			}
	
	@Test
	public void testUsingExcel() throws Exception {
		String username = ExcelUtility_SingleDataSet.getCellData(1, 0);
		String password = ExcelUtility_SingleDataSet.getCellData(1, 1);
				//Click login button
				driver.findElement(By.xpath("//div[@id='navbar']//a[contains(text(),'Login')]")).click();
				Thread.sleep(2000);
				//Enter username
				driver.finElement(By.id("user_email")).sendKeys(username);
				
				//Enter password
				driver.finElement(By.id("user_password")).sendKeys(password);

				//Click login button
				driver.finElement(By.name("commit")).click();
				
				Thread.sleep(2000);
				
				//Find if error messages exist
				boolean result = driver.findElements(By.xpath("//form[@id='new_user']//div[3]")).size() !=0;
				Assert.assertTrue(result);
				
				//Set the test result in the Excel file
				ExcelUtility_SingleDataSet.setCellData("Pass", 1, 2);
			}
	
	@AfterClass
	public void tearDown() throws Exception {
		driver.quit();
	}

}
