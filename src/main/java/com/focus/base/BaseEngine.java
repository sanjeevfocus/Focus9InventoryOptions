package com.focus.base;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;
import org.testng.log4testng.Logger;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import com.focus.Pages.LoginPage;
import com.focus.utilities.DriverUtility;
import com.focus.utilities.ScreenshotUtility;

public class BaseEngine {

	private static WebDriver driver;
	private static String baseDir;
	private static String tcName;
	private static ExtentReports extentReports;
	private static ExtentTest extentTest;
	private static Logger logger;
	private static WebDriverWait wait;

	private String filePath;
	private FileInputStream fip;
	private FileOutputStream fop;
	private Workbook workbook;
	private Sheet sheet;
	private Cell cell;
	private CellStyle style;
	private Row row;
	private String res = null;

	private static Alert alert;

	// private String xlfile="\\E:\\ZFocusAI\\TestData.xls";

	@Parameters("browser")
	@BeforeSuite
	public void openBrowser(@Optional("chrome") String browser) {

		baseDir = System.getProperty("user.dir");
		PropertyConfigurator.configure(baseDir + "\\log4j.properties");

		if (browser.equalsIgnoreCase("chrome")) {
			System.setProperty(DriverUtility.chromeKey, baseDir + DriverUtility.chromeValue);
			String downloadFilepath = getBaseDir() + "\\autoIt\\ExportFiles";
			HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
			chromePrefs.put("profile.default_content_settings.popups", 0);
			chromePrefs.put("download.default_directory", downloadFilepath);
			chromePrefs.put("safebrowsing.enabled", "true");
			ChromeOptions options = new ChromeOptions();
			options.setExperimentalOption("prefs", chromePrefs);
			DesiredCapabilities cap = DesiredCapabilities.chrome();
			cap.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
			cap.setCapability(ChromeOptions.CAPABILITY, options);

			driver = new ChromeDriver(cap);
			initActivities();
		}

		/*
		 * if(browser.equalsIgnoreCase("chrome")) {
		 * System.setProperty(DriverUtility.chromeKey,
		 * baseDir+DriverUtility.chromeValue); driver=new ChromeDriver();
		 * initActivities(); }
		 */

		if (browser.equalsIgnoreCase("ie")) {
			System.setProperty(DriverUtility.ieKey, baseDir + DriverUtility.ieValue);
			driver = new InternetExplorerDriver();
			initActivities();
		}

		if (browser.equalsIgnoreCase("firefox")) {
			System.setProperty(DriverUtility.firefoxKey, baseDir + DriverUtility.firefoxValue);
			driver = new FirefoxDriver();
			initActivities();
		}
	}

	public void initActivities() {
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		driver.manage().deleteAllCookies();
	}

	@BeforeMethod
	public void beforeMethodExecution(Method method) {
		tcName = method.getName();
		// System.out.println("tc started");
		logger = Logger.getLogger(BaseEngine.class);

		getLogger().info("Currently Executing Test Case Name is : " + tcName);
		extentTest = extentReports.startTest(tcName);

		System.out.println("This Is Before Method Execution And Trying To Executing The Method: " + tcName);
		// Excel Update is Written as below
	}

	@AfterMethod
	public void afterMethodExecution(ITestResult result)
			throws IOException, EncryptedDocumentException, InvalidFormatException {
		// System.out.println(res);
		// System.out.println(result.getStatus());
		// System.out.println("This is After method");

		if (result.getStatus() == ITestResult.SUCCESS) {
			res = "Pass";
			getLogger().info("Test Case is Passed : " + tcName);
			// excelReader.setCellData(xlfile, "Sheet3", 0, 0, res);
			extentTest.log(LogStatus.PASS, "Test Case is Passed  : " + tcName);
			System.out.println("Pass : Test Case is Passed from after Method : " + tcName);
		}

		else if (result.getStatus() == ITestResult.FAILURE) {
			res = "Fail";
			getLogger().info("Test Case is Failed : " + tcName + "So Taking the Screenshot");
			// excelReader.setCellData(xlfile, "Sheet3", 0, 1, res);
			extentTest.log(LogStatus.FAIL, "Test Case is Failed : " + tcName);
			ScreenshotUtility.screenshot();
			System.out.println("Fail : Test Case Is Failed After Method Execution : " + tcName);
			// System.out.println("Log Message:: @AfterMethod:
			// Method-"+methodName+"- has
			// Failed");

		} else if (result.getStatus() == ITestResult.SKIP) {
			res = "Skip";
			getLogger().info("Test Case is Skipped : " + tcName + "So Taking the Screenshot");
			// excelReader.setCellData(xlfile, "Sheet3", 0, 2, res);
			extentTest.log(LogStatus.SKIP, "Test Case is Skipped : " + tcName);
			ScreenshotUtility.screenshot();
			System.out.println("Skip : Test Case Is Skipped After Method Execution : " + tcName);
		}

	}

	@AfterSuite
	public void closeBrowser() {
		if (driver == null) {
			getLogger().info("Driver is pointing to Null : @AfterSuite");
		} else {
			driver.close();
		}

	}

	@BeforeTest
	public void intialiseReports() {
		extentReports = new ExtentReports(baseDir + "\\reports\\FocusAutomationReport.html");
		extentReports.assignProject("FocusAI");
	}

	@AfterTest
	public void generateReports() {
		extentReports.flush();
		extentReports.endTest(extentTest);
	}

	public static String getTCName() {
		return tcName;
	}

	public static String getBaseDir() {
		return baseDir;
	}

	public static WebDriver getDriver() {
		return driver;
	}

	public static ExtentTest getExtentTest() {
		return extentTest;
	}

	public static Logger getLogger() {
		return logger;
	}

	public static void enterUrl(String url) {
		getDriver().get(url);
	}

	public static boolean isElementNotDisplayed(WebElement we) {
		try {
			if (we.isDisplayed()) {
				return true;
			}

			return false;
		} catch (Exception e) {
			return false;
		}
	}

	public static Actions getAction() {
		Actions action = new Actions(driver);
		return action;

	}

	public static void getWaitForAlert() {
		new WebDriverWait(getDriver(), 800).ignoring(NoAlertPresentException.class)
				.until(ExpectedConditions.alertIsPresent());
	}

	public static Alert getAlert() {
		alert = getDriver().switchTo().alert();
		return alert;
	}

	public static void getFrame(WebElement ele) {
		getDriver().switchTo().frame(ele);

	}

	public static boolean getIsAlertPresent() {
		try {
			driver.switchTo().alert();
			return true;
		} catch (NoAlertPresentException Ex) {
			return false;
		}
	}

	public static Wait getWebDriverWait() {
		wait = new WebDriverWait(getDriver(), 90);
		return wait;

	}

	public static Wait getFluentWebDriverWait() {
		Wait<WebDriver> wait = new FluentWait<WebDriver>(driver).withTimeout(40, TimeUnit.SECONDS)
				.pollingEvery(5, TimeUnit.SECONDS).ignoring(NoSuchElementException.class);
		return wait;
	}

	public static String getDate() {

		DateFormat dateFormat = new SimpleDateFormat("MMMM yyyy");
		Calendar cal = Calendar.getInstance();
		// System.out.println(dateFormat.format(cal.getTime()));
		return dateFormat.format(cal.getTime());

	}

	public static String getCurrentDate() {
		Date date = Calendar.getInstance().getTime();
		DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
		String strDate = dateFormat.format(date);
		// System.out.println(strDate);
		return strDate;

	}

	public static String getCurrentDateF2() {
		Date date = Calendar.getInstance().getTime();
		DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
		String strDate = dateFormat.format(date);
		// System.out.println(strDate);
		return strDate;

	}

	public static String getCurrentDateFormat() {
		Date date = Calendar.getInstance().getTime();
		DateFormat dateFormat = new SimpleDateFormat("M/d/yyyy");
		String strDate = dateFormat.format(date);
		// System.out.println(strDate);
		return strDate;

	}

	public static String getTextByJs(WebElement element) { // for only getText (Individual text only)
		
		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		String text = (String) js.executeScript("return arguments[0].innerHTML", element);
		return text;
	}

	public static void ClickUsingJs(WebElement element) {

		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].click();", element);

	}

	public static void click(WebElement element) {

		new WebDriverWait(getDriver(), 50).until(ExpectedConditions.elementToBeClickable(element));

		element.click();
	}

	public static void enterText(WebElement element, String Value) {

		new WebDriverWait(getDriver(), 50).until(ExpectedConditions.elementToBeClickable(element));

		element.sendKeys(Value);
	}

	public static String getText(WebElement element) {

		new WebDriverWait(getDriver(), 50).until(ExpectedConditions.elementToBeClickable(element));

		return element.getText();
	}

	public static void clearValueFromTextBox(WebElement element) {

		new WebDriverWait(getDriver(), 50).until(ExpectedConditions.elementToBeClickable(element));

		element.clear();
	}

	public static void tab(WebElement element) {

		new WebDriverWait(getDriver(), 50).until(ExpectedConditions.elementToBeClickable(element));

		element.sendKeys(Keys.TAB);
	}

	public static void ScrollToElement(WebElement element) {

		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView();", element);
	}

	public static String getMonth() {
		DateFormat dateFormat = new SimpleDateFormat("MMMM");
		Calendar cal = Calendar.getInstance();
		// System.out.println(dateFormat.format(cal.getTime()));
		return dateFormat.format(cal.getTime());
	}

	public static void dropDown(WebElement element, String value) {

		new Select(element).selectByValue(value);

	}

	public static void dropDown(WebElement element, int index) {

		new Select(element).selectByIndex(index);

	}

	public static void dropDown(String str, WebElement element) {

		new Select(element).selectByVisibleText(str);

	}

	public void contextClick(WebElement element) {

		try {
			Actions action = new Actions(getDriver()).contextClick(element);

			action.build().perform();

			System.out.println("right clicked the element");
		} catch (StaleElementReferenceException e) {
			System.out.println("Element is not attached to the page document " + e.getStackTrace());
		} catch (NoSuchElementException e) {
			System.out.println("Element " + element + " was not found in DOM " + e.getStackTrace());
		} catch (Exception e) {
			System.out.println("Element " + element + " was not clickable " + e.getStackTrace());
		}
	}

	public void moveToElement(WebElement element) {
		try {
			Actions action = new Actions(getDriver()).moveToElement(element).click();

			action.build().perform();

			System.out.println("MoveTo element clicked the element");
		} catch (StaleElementReferenceException e) {
			System.out.println("Element is not attached to the page document " + e.getStackTrace());
		} catch (NoSuchElementException e) {
			System.out.println("Element " + element + " was not found in DOM " + e.getStackTrace());
		} catch (Exception e) {
			System.out.println("Element " + element + " was not clickable " + e.getStackTrace());
		}
	}
	
	
	@FindBy(xpath="//span[@class='hidden-xs']")
	public static WebElement userNameDisplay;
	
	@FindBy(xpath="//*[@id='companyLogo']")
	public static WebElement companyLogo;

	@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
	public static WebElement companyName;

	@FindBy(xpath="//*[@id='userprofile']/li/span[2]")
	public static WebElement logoutOption;
	
	
	
	 
	 @FindBy(xpath = "//*[@id='errmsgDiv']")
		public static WebElement loginPageErrorMesg;

		@FindBy(id = "txtPassword")
		public static WebElement password;

		@FindBy(id = "ddlCompany")
		public static WebElement companyDropDownList;

		@FindBy(id = "btnSignin")
		public static WebElement signIn;
		
		@FindBy(xpath="//*[@id='txtUsername']")
		private static WebElement username;
	
	
public void reLogin(String userName, String passWord) throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		
		//excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Thread.sleep(3000);
		click(userNameDisplay);
		Thread.sleep(1000);
		click(logoutOption);
		Thread.sleep(3000);

		//String unamelt = excelReader.getCellData(xlSheetName, 102, 5);

		//String pawslt = excelReader.getCellData(xlSheetName, 103, 5);

		LoginPage.enterUserName(userName);
		LoginPage.enterPassword(passWord);

		Select company = new Select(companyDropDownList);
		company.selectByValue("108");
		Thread.sleep(2000);

		LoginPage.clickOnSignInBtn();
		
		/*try {
			
			if(loginPageErrorMesg.isDisplayed())
			{
				
				enterText(username, userName);
				enterText(password, passWord);
				Select company1 = new Select(companyDropDownList);
				company1.selectByValue("108");
				Thread.sleep(2000);
				click(signIn);
			}
			
		} catch (Exception e) {
			
			System.err.println("Error Message NOT Found or NOT Clickable");
		System.err.println(e.getMessage());

		String Exception = e.getMessage();

		//return Exception;
			
			
		}*/
	
		
		

		Thread.sleep(3000);

		System.out.println("Signing");
	}

@FindBy (xpath="//*[@id='25']")
public static WebElement dataMangementMenu;

@FindBy (xpath="//*[@id='27']/span")
private static WebElement restore;

@FindBy (xpath="//button[@class='btn Fbutton']")
private static WebElement folderpathExpandBtn;


@FindBy (xpath="//*[@class='icon-restore hiconright2']")
private static WebElement restoreCompanyBtn;

@FindBy (xpath="//*[@id='RestoreCompanyModal']/div/div/div[2]/button[1]")
private static WebElement overRideYesBtn;

@FindBy (xpath="//a[@id='1']")
public static WebElement homeMenu;

@FindBy(xpath="//*[@id='id_mainlayoutmenu']/ul[2]/li[6]")
	public static WebElement LogoutDropdown;




public static  void restoreCompany(String BackupFile, String companyName) throws InterruptedException, IOException, AWTException
{
	
	Thread.sleep(1000);

	getDriver().navigate().refresh();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
	dataMangementMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restore));
	restore.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(folderpathExpandBtn));
	folderpathExpandBtn.click();
	
	Thread.sleep(3000);
	
	 Robot rb = new Robot();
     StringSelection str = new StringSelection(getBaseDir()+"\\backup\\"+BackupFile+".fbak");
     Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str, null);
 
     rb.keyPress(KeyEvent.VK_CONTROL);
     rb.keyPress(KeyEvent.VK_V);
 
    rb.keyRelease(KeyEvent.VK_CONTROL);
    rb.keyRelease(KeyEvent.VK_V);
 
    rb.keyPress(KeyEvent.VK_ENTER);
    rb.keyRelease(KeyEvent.VK_ENTER);

     
     Thread.sleep(3000);
     
     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtn));
     restoreCompanyBtn.click();
     
     Thread.sleep(3000);
     
     try
     {
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(overRideYesBtn));
 		 overRideYesBtn.click();
 		 
 		 System.err.println("COMPANY EXISTS");
 		 
 		Thread.sleep(130000);
 		 
     }
     catch(Exception e)
     {
    	 System.err.println("NO OLDER COMPANY EXISTS");
     }
     
     if(getIsAlertPresent())
     {
    	 System.err.println("Alert Displayed");
    	 getWaitForAlert();
    	 System.err.println(getAlert().getText());
    	 getAlert().accept();
     }
     
     Thread.sleep(3000);
	
     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LogoutDropdown));
     LogoutDropdown.click();
	
     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
     logoutOption.click();

     Thread.sleep(3000);
     
     checkLoginToSelectedCompany(companyName,"su","su");
     
     Thread.sleep(30000);
     

     getDriver().navigate().refresh();
     
     Thread.sleep(3000);
     
     
     checkLoginToSelectedCompany(companyName,"su","su");
   
}




public static void checkLoginToSelectedCompany(String cname,String username,String Password) throws InterruptedException
{
	  	Thread.sleep(3000);
	     
	     LoginPage lp= new LoginPage(getDriver());
	     
	     lp.enterUserName(username);
	     
	     lp.enterPassword(Password);
	     
			String compname = cname;

			Select oSelect = new Select(companyDropDownList);

			List<WebElement> elementCount = oSelect.getOptions();

			int cqSize = elementCount.size();

			System.out.println("CompanyDropdownList Count :" + cqSize);

			int i;

			for (i = 0; i < elementCount.size(); i++) {

				elementCount.get(i).getText();

				String optionName = elementCount.get(i).getText();
				if (optionName.toUpperCase().startsWith(compname.toUpperCase())) {
					System.out.println("Logined Company" + elementCount.get(i).getText());
					elementCount.get(i).click();
				}

			}

			Thread.sleep(2000);

			lp.clickOnSignInBtn();
}



}
