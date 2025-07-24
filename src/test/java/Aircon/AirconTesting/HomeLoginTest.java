package Aircon.AirconTesting;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class HomeLoginTest {
	public static WebDriver driver;
	public static String path = "C:\\Users\\Hi\\Desktop\\excel\\AirconTrip B2B.xlsx";
	public static FileInputStream fs;
	public static FileOutputStream fos;
	public static Workbook wb;
	public static Sheet sheet1;

	@BeforeMethod
	public void setUp() throws IOException {
	    ChromeOptions options = new ChromeOptions();
	    options.addArguments("--headless"); // run in headless mode
	    driver = new ChromeDriver(options); // use class-level driver
	    driver.manage().window().maximize();
	

		fs = new FileInputStream(path);
		wb = new XSSFWorkbook(fs);
		sheet1 = wb.getSheetAt(0);
	}

	@Test
	public void TestCase01() throws InterruptedException {
		Row row = sheet1.getRow(1);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-01")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//a[contains(@href, '/why-my-partner')]")).click();
			Thread.sleep(2000);
			String url = driver.getCurrentUrl();

			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (url.equals("https://business.aircontrip.com/why-my-partner")) {
				resultCell.setCellValue("Partner Page Opened");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(greenStyle);
			} else {
				resultCell.setCellValue("Other page loaded");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(redStyle);
			}
		}
	}

	@Test
	public void TestCase02() throws InterruptedException {
		Row row = sheet1.getRow(2);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-02")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//a[contains(@href, '/product-highlights')]")).click();
			Thread.sleep(2000);
			String url = driver.getCurrentUrl();

			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (url.equals("https://business.aircontrip.com/product-highlights")) {
				resultCell.setCellValue("Highlights Opened");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(greenStyle);
			} else {
				resultCell.setCellValue("Other page loaded");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(redStyle);
			}
		}
	}

	@Test
	public void TestCase03() throws InterruptedException {
		Row row = sheet1.getRow(3);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-03")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);

			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			try {
				WebElement svgButton = driver
						.findElement(By.xpath("//button//svg[contains(@xmlns, 'http://www.w3.org/2000/svg')]"));
				svgButton.click();
				resultCell.setCellValue("Working");
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} catch (Exception e) {
				System.out.println("SVG button not found or not clickable: " + e.getMessage());
				resultCell.setCellValue("Not Working");
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}
	}

	@Test
	public void TestCase05() throws InterruptedException {
		Row row = sheet1.getRow(5);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-05")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//button[@type='submit']")).click();

			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (driver.getCurrentUrl().equals("https://business.aircontrip.com/agent/flight/dashboard")) {
				resultCell.setCellValue("Working on blank click");
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			} else {
				resultCell.setCellValue("Not Working on Blank Click");
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			}
		}
	}

	@Test
	public void TestCase06() throws InterruptedException {
		Row row = sheet1.getRow(6);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-06")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//input[@placeholder='Email']")).sendKeys("saquibhamza333@gmail.com");
			driver.findElement(By.xpath("//input[@placeholder='Password']")).sendKeys("pissword");
			driver.findElement(By.xpath("//button[@type='submit']")).click();
			Thread.sleep(1000);
			String popup = driver.findElement(By.xpath("//div[@role='status']")).getText();

			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			System.out.println(popup);
			if (popup.equals("Invalid email or password")) {
				resultCell.setCellValue(popup);
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}
	}

	@Test
	public void TestCase07() throws InterruptedException {
		Row row = sheet1.getRow(7);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-07")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//input[@placeholder='Email']")).sendKeys("saquibhamza@gmail.com");
			driver.findElement(By.xpath("//input[@placeholder='Password']")).sendKeys("password");
			driver.findElement(By.xpath("//button[@type='submit']")).click();
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			WebElement popupElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@role='status']")));
			String popup = popupElement.getText();

			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			System.out.println(popup);
			if (popup.equals("Invalid email or password")) {
				resultCell.setCellValue(popup);
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}
	}

	@Test
	public void TestCase08() throws InterruptedException {
		Row row = sheet1.getRow(8);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-08")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//input[@placeholder='Email']")).sendKeys("saquibhamza333@gmail.com");
			driver.findElement(By.xpath("//input[@placeholder='Password']")).sendKeys("password");
			driver.findElement(By.xpath("//button[@type='submit']")).click();
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			WebElement popupElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@role='status']")));
			String popup = popupElement.getText();
			
			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			System.out.println(popup);
			if (popup.equals("Login successful")
					&& driver.getCurrentUrl().equals("https://business.aircontrip.com/agent/flight/dashboard")) {
				resultCell.setCellValue(popup);
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				resultCell.setCellValue(popup);
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}

		}
	}

	@Test
	public void TestCase09() throws InterruptedException {
		Row row = sheet1.getRow(9);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-09")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//input[@placeholder='Password']")).sendKeys("Hello");

			driver.findElement(By.xpath("//button[@aria-label='Show password']//*[name()='svg']")).click();
			Thread.sleep(1000);
			String hidden = driver.findElement(By.xpath("//input[@placeholder='Password']")).getAttribute("type");
			;

			Thread.sleep(2000);
			driver.findElement(By.xpath("//button[@aria-label='Hide password']//*[name()='svg']")).click();
			Thread.sleep(1000);
			String visible = driver.findElement(By.xpath("//input[@placeholder='Password']")).getAttribute("type");
			;
			System.out.println(hidden);
			System.out.println(visible);

			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (hidden.equals("password") && visible.equals("text")) {
				resultCell.setCellValue("Eye View working fine but will work vice versa");
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				resultCell.setCellValue("Eye View not working");
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}
	}

	@Test
	public void TestCase10() throws InterruptedException {
		Row row = sheet1.getRow(10);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-10")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//input[@placeholder='Email']")).sendKeys("saquibhamza@fg");
			driver.findElement(By.xpath("//input[@placeholder='Password']")).sendKeys("Hello");
			driver.findElement(By.xpath("//button[@type='submit']")).click();
			Thread.sleep(1000);
			String pop = driver.findElement(By.xpath(
					"//input[@type='email' and @placeholder='Email']/ancestor::div[contains(@class, 'flex-col')]/p"))
					.getText();
			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (pop.equals("Invalid email address")) {
				resultCell.setCellValue(pop);
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				resultCell.setCellValue(pop);
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}
	}

	@Test
	public void TestCase11() throws InterruptedException {
		Row row = sheet1.getRow(11);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-11")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//a[@href='/forgot-password']")).click();
			Thread.sleep(5000);
			String url = driver.getCurrentUrl();
			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (url.equals("https://business.aircontrip.com/forgot-password")) {
				resultCell.setCellValue(url);
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				resultCell.setCellValue(url);
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}

	}

	@Test
	public void TestCase12() throws InterruptedException {
		Row row = sheet1.getRow(12);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-12")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//a[@href='/forgot-password']")).click();
			Thread.sleep(5000);
			WebElement emailInput = driver.findElement(By.xpath("//input[@type='email']"));
			String requiredAttr = emailInput.getAttribute("required");

			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (requiredAttr != null) {
				resultCell.setCellValue("Required email is present");
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				resultCell.setCellValue("required email is not present");
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}
	}

	@Test
	public void TestCase13() throws InterruptedException {
		Row row = sheet1.getRow(13);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-13")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//a[@href='/forgot-password']")).click();
			Thread.sleep(5000);
			driver.findElement(By.xpath("//input[@type='email']")).sendKeys("faizan@gmail.com");
			driver.findElement(By.xpath("//button[@type='submit']")).click();
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
			WebElement popup = wait
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@role='status']")));

			String pop = popup.getText();
			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (pop.equals("Admin not found")) {
				resultCell.setCellValue(pop);
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				resultCell.setCellValue(pop);
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}

		}
	}

	@Test
	public void TestCase14() throws InterruptedException {
		Row row = sheet1.getRow(14);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-14")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//a[@href='/forgot-password']")).click();
			Thread.sleep(5000);
			driver.findElement(By.xpath("//input[@type='email']")).sendKeys("faizan@gmail");
			driver.findElement(By.xpath("//button[@type='submit']")).click();
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			WebElement popup = wait
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@role='status']")));

			String pop = popup.getText();
			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (pop.equals("Enter Email Format")) {
				resultCell.setCellValue(pop);
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				resultCell.setCellValue(pop);
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}
	}

	@Test
	public void TestCase15() throws InterruptedException {
		Row row = sheet1.getRow(15);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-15")) {
			driver.get("https://business.aircontrip.com");
			Thread.sleep(3000);
			driver.findElement(By.xpath("//a[@href='/forgot-password']")).click();
			Thread.sleep(5000);
			driver.findElement(By.xpath("//input[@type='email']")).sendKeys("saquibhamza333@gmail.com");
			driver.findElement(By.xpath("//button[@type='submit']")).click();
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			WebElement popup = wait
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@role='status']")));

			String pop = popup.getText();
			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (pop.equals("Reset link sent to your email")) {
				resultCell.setCellValue(pop);
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				resultCell.setCellValue(pop);
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}
	}

	@Test
	public void TestCase25() throws InterruptedException {
		Row row = sheet1.getRow(25);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-25")) {
			driver.get("https://business.aircontrip.com");
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
			String initialUrl = driver.getCurrentUrl();
			Thread.sleep(5000);
			driver.findElement(By.xpath("//button[text()='GET STARTED']")).click();
			Thread.sleep(2000);
			String afterClickUrl = driver.getCurrentUrl();

			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			// Check if URL changed
			if (!afterClickUrl.equals(initialUrl)) {
				resultCell.setCellValue("Routing Perfectly");
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				resultCell.setCellValue("Button not working");
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}
	}

	@Test
	public void TestCase26() throws InterruptedException {
		Row row = sheet1.getRow(26);
		Cell cell = row.getCell(0);
		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-26")) {
			driver.get("https://business.aircontrip.com");
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
			String initialUrl = driver.getCurrentUrl();
			Thread.sleep(5000);
			driver.findElement(By.xpath("//button[text()='CONTACT US']")).click();
			Thread.sleep(2000);
			String afterClickUrl = driver.getCurrentUrl();

			Cell resultCell = row.createCell(7);
			Cell statusCell = row.createCell(8);

			CellStyle greenStyle = wb.createCellStyle();
			greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle redStyle = wb.createCellStyle();
			redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			// Check if URL changed
			if (!afterClickUrl.equals(initialUrl)) {
				resultCell.setCellValue(afterClickUrl);
				statusCell.setCellStyle(greenStyle);
				statusCell.setCellValue("Passed");
			} else {
				resultCell.setCellValue("Button not working");
				statusCell.setCellStyle(redStyle);
				statusCell.setCellValue("Failed");
			}
		}
	}

	@AfterMethod
	public void tearDown() {
		try {
			if (fs != null)
				fs.close();

			FileOutputStream fos = new FileOutputStream(path);
			if (wb != null) {
				wb.write(fos);
				wb.close();
			}
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (driver != null) {
				System.out.println("Closing browser...");
				driver.quit();
			}
		}
	}

}
