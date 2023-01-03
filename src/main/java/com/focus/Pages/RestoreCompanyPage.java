package com.focus.Pages;


import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;


public class RestoreCompanyPage extends BaseEngine
{

	//Restore Button in the Login Page
	@FindBy(xpath="//i[@class='icon-restore-1 icon-font4']")
	private static WebElement  loginPageRestoreBtn;
	
	//User Name and password in Credentials
	@FindBy(xpath="//*[@id='txtAdminUsername']")
	private static WebElement adminName;
	
	@FindBy(xpath="//*[@id='txtAdminPassword']")
	private static WebElement adminPassword;
	
	@FindBy(xpath="/html/body/section/div[8]/div/div/div[4]/button[1]")
	private static WebElement loginButton;
	
	@FindBy(xpath="/html/body/section/div[8]/div/div/div[4]/button[2]")
	private static WebElement cancelButton;

	//Restore Screen Options
	@FindBy(xpath="//div[contains(text(),'Restore Company')]")
	private static WebElement restoreCompanyBtnOfLoginPage;
	
	@FindBy(xpath="//input[@id='txtRestoreFile']")
	private static WebElement restoreInputTxt;
	
	@FindBy(xpath="//button[@class='btn Fbutton']")
	private static WebElement restoreBrowseBtn;

    @FindBy(id="ddlCompany")
    private static WebElement companyDropDownList;
	
	private static ExcelReader excelReader;
	
	private static String xlfile;
	private static String resPass="Pass";
	private static String resFail="Fail";
		
	public static int actFirstCompanyCodeIncreament=0;
	
	public static int actCompanyCodeIncreament=0;
	
	public static String CompCode;
	
	public static String actCompanyWithCodeOnly;
	
	public static String actAppendCompanyCode;
	
	 // Click on Restore Button   
	 public boolean checkAdminCredentialsScreenOnClickOnRestoreOptionInLoginPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyDropDownList));
		companyDropDownList.click();
		Thread.sleep(2000);
		companyDropDownList.sendKeys(Keys.END);
		Thread.sleep(2000);
		companyDropDownList.sendKeys(Keys.ENTER);
		Thread.sleep(2000);
		
		String actCompanyName=companyDropDownList.getText();
		
		System.out.println("actCompanyName"+actCompanyName);
	
		String[] lines = actCompanyName.split("\n");

		System.out.println("First Company");
	
		String firstLine=lines[0];
        
        String actFirstCompanyCode = "";
        
        String firstCompCode=firstLine.substring(firstLine.length()-4);
        
        actFirstCompanyCode = firstCompCode.substring(1, 2);
        
        System.out.println("actFirstCompanyCode"+actFirstCompanyCode);
		
        System.out.println("Last Company");
			
        String lastLine=lines[lines.length-2];
        
        String actCompanyCode = "";
        
        
        String actCompanyWithCodeOnly="";
        
        String CompCode=lastLine.substring(lastLine.length()-4);
        
        actCompanyCode = CompCode.substring(1, 2);
        
        actCompanyWithCodeOnly=CompCode.substring(0, 3);
       
        System.out.println("actCompanyCode"+actCompanyCode);
       
        
        //Checking the Company Code is Numeric or Alphabets
        
        if ((Character.isLetter(actCompanyCode.charAt(0)) == false)) 
        {
            System.out.println("Output is Numeric");
            
            int actCompanyCodeConvertToInt=Integer.parseInt(actFirstCompanyCode);
            
            actFirstCompanyCodeIncreament=actCompanyCodeConvertToInt;
            
         }
        else
        {
        	System.out.println("Output is Alphabets");
        	
        }
        
        
        if ((Character.isLetter(actCompanyCode.charAt(0)) == false)) 
        {
            System.out.println("Output is Numeric");
            
            int actCompanyCodeConvertToInt=Integer.parseInt(actCompanyCode);
            
            actCompanyCodeIncreament=actCompanyCodeConvertToInt+1;
            
         }
        else
        {
        	System.out.println("Output is Alphabets");
        	
        }
        
            
        System.out.println("actCompanyCodeIncreament"+actCompanyCodeIncreament);
	
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginPageRestoreBtn));
		loginPageRestoreBtn.click();
		
		getWebDriverWait().until(ExpectedConditions.visibilityOf(adminName));
    	
        boolean actAdminUserNameInRestore  =adminName.isDisplayed();
        boolean actAdminPasswordInRestore  =adminPassword.isDisplayed();
        boolean actAdminLoginBtnInRestore  =loginButton.isDisplayed();
        boolean actAdminCancelBtnInRestore =cancelButton.isDisplayed();
        
        boolean expAdminUserNameInRestore  =true;
        boolean expAdminPasswordInRestore  =true;
        boolean expAdminLoginBtnInRestore  =true;
        boolean expAdminCancelBtnInRestore =true;
                
        System.out.println("********************************** checkAdminCredentialsScreenOnClickOnRestoreOptionInLoginPage ********************************************");
		System.out.println("Admin Screen User Name Availablility Option Value Actual     :" + actAdminUserNameInRestore   +"  Value Expected :" + expAdminUserNameInRestore);
		System.out.println("Admin Screen Password Availablility Option Value Actual      :" + actAdminPasswordInRestore   +"  Value Expected :" + expAdminPasswordInRestore);
		System.out.println("Admin Screen Ok Button Availablility Option Value Actual     :" + actAdminLoginBtnInRestore   +"  Value Expected :" + expAdminLoginBtnInRestore);
		System.out.println("Admin Screen Cancel Button Availablility Option Value Actual :" + actAdminCancelBtnInRestore  +"  Value Expected :" + expAdminCancelBtnInRestore);
    
	    if(actAdminUserNameInRestore==expAdminUserNameInRestore && actAdminPasswordInRestore==expAdminPasswordInRestore
	    		&& actAdminLoginBtnInRestore==expAdminLoginBtnInRestore && actAdminCancelBtnInRestore==expAdminCancelBtnInRestore)
	    {
		    excelReader.setCellData(xlfile, "Sheet1", 13, 9, resPass);
			return true;
        }
	    else 
	    {
			excelReader.setCellData(xlfile, "Sheet1", 13, 9, resFail);
			return false;
		
		 }
	 }
	 



	 public static boolean checkLoginButtonInAdminCredentialsByInputUserNameAndPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(adminName));
		adminPassword.sendKeys("admin");
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginButton));
		loginButton.click();
		
		Thread.sleep(2000);
		
        boolean actRestoreCompanyBtnOfLoginPage  =restoreCompanyBtnOfLoginPage.isDisplayed();
        boolean actRestoreInputTxt               =restoreInputTxt.isDisplayed();
        boolean actRestoreBrowseBtn              =restoreBrowseBtn.isDisplayed();
        
        boolean expRestoreCompanyBtnOfLoginPage  =true;
        boolean expRestoreInputTxt               =true;
        boolean expRestoreBrowseBtn              =true;
        
                
        System.out.println("********************************** checkAdminCredentialsScreenOnClickOnRestoreOptionInLoginPage ********************************************");
		System.out.println("Admin Screen User Name Availablility Option Value Actual    :"+actRestoreCompanyBtnOfLoginPage  +"  Value Expected :"+expRestoreCompanyBtnOfLoginPage);
		System.out.println("Admin Screen Password Availablility Option Value Actual     :"+actRestoreInputTxt               +"  Value Expected :"+expRestoreInputTxt);
		System.out.println("Admin Screen Ok Button Availablility Option Value Actual    :"+actRestoreBrowseBtn              +"  Value Expected :"+expRestoreBrowseBtn);
		
    
	    if(actRestoreCompanyBtnOfLoginPage==expRestoreCompanyBtnOfLoginPage && actRestoreInputTxt==expRestoreInputTxt
	    		&& actRestoreBrowseBtn==expRestoreBrowseBtn)
	    {
		    excelReader.setCellData(xlfile, "Sheet1", 13, 9, resPass);
			return true;
        }
	    else 
	    {
			excelReader.setCellData(xlfile, "Sheet1", 13, 9, resFail);
			return false;
		
		 }
	 }
	 
	 
	 @FindBy(xpath="//div[@class='theme_color font-6']")
	 public static WebElement errorMessage;
		 
	 @FindBy(xpath="//span[@class='icon-reject2 theme_color']")
	 public static WebElement errorMessageCloseBtn;
	 
	 
	 
	 public static boolean checkClickOnRestoreCompanyWithoutPath() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtnOfLoginPage));
		restoreCompanyBtnOfLoginPage.click();
		

		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
		String actGetValidationMessageOnClickOnRestoreWithBlankPath=errorMessage.getText();		
		String expGetValidationMessageOnClickOnRestoreWithBlankPath="Please select BackUp File";
		errorMessageCloseBtn.click();
                
        System.out.println("*********************************************************************** checkClickOnRestoreCompanyWithoutPath *************************************************************************************************");
		System.out.println("Restore Company Screen Click on Restore Company without Path Value Actual    :"+actGetValidationMessageOnClickOnRestoreWithBlankPath  +"  Value Expected :"+expGetValidationMessageOnClickOnRestoreWithBlankPath);

		
    
	    if(actGetValidationMessageOnClickOnRestoreWithBlankPath.equalsIgnoreCase(expGetValidationMessageOnClickOnRestoreWithBlankPath))
	    {
		    excelReader.setCellData(xlfile, "Sheet1", 13, 9, resPass);
			return true;
        }
	    else 
	    {
			excelReader.setCellData(xlfile, "Sheet1", 13, 9, resFail);
			return false;
		
		 }
	 }
	 
	 
	 
	 @FindBy(xpath="//input[@id='txtRestoreFile']")
	 public static WebElement restoreCompanyNameTxt;
	 
	 @FindBy(xpath="//*[@id='dvCompanyDetails']/div[2]/label")
	 public static WebElement restoreFileName;
	 
	 @FindBy(xpath="//*[@id='dvCompanyDetails']/div[4]/label")
	 public static WebElement restoreCompanyName;
	 
	 @FindBy(xpath="//*[@id='dvCompanyDetails']/div[6]/label")
	 public static WebElement restoreAccountingDate;
	 
	 @FindBy(xpath="//*[@id='dvCompanyDetails']/div[8]/label")
	 public static WebElement restoreBackUpDate;
	 
	 @FindBy(xpath="//*[@id='dvCompanyDetails']/div[10]/label")
	 public static WebElement restoreCreatedBy;
	 
	 public static boolean checkClickOnBrowseBtnAndSelectBackupInRestoreCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreBrowseBtn));
		restoreBrowseBtn.click();
		
		Thread.sleep(5000);
		
		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\backupSelectionInRestoreCompany.exe");
		
		Thread.sleep(5000);
		
		String actRestoreCompanyNameTxt =restoreCompanyNameTxt.getAttribute("value");
		String actRestoreFileName       =restoreFileName.getText();
		String actRestoreCompanyName    =restoreCompanyName.getText();
		String actRestoreAccountingDate =restoreAccountingDate.getText();
		String actRestoreBackupDate     =restoreBackUpDate.getText();
		String actRestoreCreatedBy      =restoreCreatedBy.getText();
		
		
		
		String expRestoreCompanyNameTxt="Automation Company.fbak";
		String expRestoreFileName      =":Automation Company.fbak";
		String expRestoreCompanyName   =":Automation Company";
		String expRestoreAccountingDate=":1/1/2019";
		String expRestoreBackUpDate    =":12/4/2019 10:45:34 AM";
		String expRestoreCreatedBy     =":SU";
		
		
                
        System.out.println("**************************************** checkClickOnBrowseBtnAndSelectBackupInRestoreCompany ***********************************************");
		System.out.println("Restore Company Screen after Selecting Backup Value Actual    :"+ actRestoreCompanyNameTxt  +"  Value Expected :"+ expRestoreCompanyNameTxt);
		System.out.println("Restore Company Screen after Selecting Backup Value Actual    :"+ actRestoreFileName        +"  Value Expected :"+ expRestoreFileName);
		System.out.println("Restore Company Screen after Selecting Backup Value Actual    :"+ actRestoreCompanyName     +"  Value Expected :"+ expRestoreCompanyName);
		System.out.println("Restore Company Screen after Selecting Backup Value Actual    :"+ actRestoreAccountingDate  +"  Value Expected :"+ expRestoreAccountingDate);
		System.out.println("Restore Company Screen after Selecting Backup Value Actual    :"+ actRestoreBackupDate      +"  Value Expected :"+ expRestoreBackUpDate);
		System.out.println("Restore Company Screen after Selecting Backup Value Actual    :"+ actRestoreCreatedBy       +"  Value Expected :"+ expRestoreCreatedBy);
		
    
	    if(actRestoreCompanyNameTxt.equalsIgnoreCase(expRestoreCompanyNameTxt) && actRestoreFileName.equalsIgnoreCase(expRestoreFileName)
	    		&& actRestoreCompanyName.equalsIgnoreCase(expRestoreCompanyName) && actRestoreAccountingDate.equalsIgnoreCase(expRestoreAccountingDate)
	    		&& actRestoreBackupDate.equalsIgnoreCase(expRestoreBackUpDate) && actRestoreCreatedBy.equalsIgnoreCase(expRestoreCreatedBy))
	    {
		    excelReader.setCellData(xlfile, "Sheet1", 13, 9, resPass);
			return true;
        }
	    else 
	    {
			excelReader.setCellData(xlfile, "Sheet1", 13, 9, resFail);
			return false;
		
		 }
	 }
	 
	 
	 @FindBy(xpath="//*[@id='restoreMSG']")
	 public static WebElement restoreCompanyMessage;
	 
	 @FindBy(xpath="//button[contains(text(),'Yes')]")
	 public static WebElement restoreCompanyMessageYes;
	 
	 @FindBy(xpath="//button[contains(text(),'No')]")
	 public static WebElement restoreCompanyMessageNo;
	 
	 @FindBy(xpath="//*[@id='RestoreCompanyModal']/div/div/div[2]/button[3]")
	 public static WebElement restoreCompanyMessageCancel;
	 

	 public static boolean checkClickOnRestoreCompanyOnSelectingPath() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtnOfLoginPage));
		restoreCompanyBtnOfLoginPage.click();
		
		Thread.sleep(5000);
			
		String actRestoreCompanyMessage        =restoreCompanyMessage.getText();
		boolean actRestoreCompanyMessageYes    =restoreCompanyMessageYes.isDisplayed();
		boolean actRestoreCompanyMessageNo     =restoreCompanyMessageNo.isDisplayed();
		boolean actRestoreCompanyMessageCancel =restoreCompanyMessageCancel.isDisplayed();

		String expRestoreCompanyMessage1       ="Company Automation Company already exists.";
	    String expRestoreCompanyMessage2       = "Do you want to overwrite existing Company?";
	    
	    boolean expRestoreCompanyMessageYes=true;
	    boolean expRestoreCompanyMessageNo=true;
	    boolean expRestoreCompanyMessageCancel=true;
	            
        System.out.println("************************************************** checkClickOnRestoreCompanyOnSelectingPath *******************************************************************************");
		System.out.println("Restore Company Screen after Selecting Backup Value Actual    :"+ actRestoreCompanyMessage       +"  Value Expected :"+ expRestoreCompanyMessage1+expRestoreCompanyMessage2);
		System.out.println("Restore Company Screen after Selecting Backup Value Actual    :"+ actRestoreCompanyMessageYes    +"  Value Expected :"+ expRestoreCompanyMessageYes);
		System.out.println("Restore Company Screen after Selecting Backup Value Actual    :"+ actRestoreCompanyMessageNo     +"  Value Expected :"+ expRestoreCompanyMessageNo);
		System.out.println("Restore Company Screen after Selecting Backup Value Actual    :"+ actRestoreCompanyMessageCancel +"  Value Expected :"+ expRestoreCompanyMessageCancel);
		
		System.out.println("actCompanyCodeIncreament"+actCompanyCodeIncreament);
		 
		System.out.println("actFirstCompanyCodeIncreament"+actFirstCompanyCodeIncreament);
		 
	    if(actRestoreCompanyMessage.startsWith(expRestoreCompanyMessage1) && actRestoreCompanyMessage.endsWith(expRestoreCompanyMessage2)
	    		&& actRestoreCompanyMessageYes==expRestoreCompanyMessageYes && actRestoreCompanyMessageNo==expRestoreCompanyMessageNo
	    		&& actRestoreCompanyMessageCancel==expRestoreCompanyMessageCancel)
	    {
		    excelReader.setCellData(xlfile, "Sheet1", 13, 9, resPass);
			return true;
        }
	    else 
	    {
			excelReader.setCellData(xlfile, "Sheet1", 13, 9, resFail);
			return false;
		
		}
	    
	 }
	 
	 
	 
	 
	    @FindBy(xpath="//*[@id='txtUsername']")
		private static WebElement username;

		@FindBy(id="txtPassword")
		private static WebElement password;

		@FindBy(id="btnSignin")
		private static WebElement signIn;
	 
	 
	 
	 public boolean checkClickOnCancelButtonInRestoreCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
		
		
        getWebDriverWait().until(ExpectedConditions.visibilityOf(restoreCompanyMessageCancel));
        restoreCompanyMessageCancel.click();
        
        Thread.sleep(5000);
        
		
		boolean actLPUserName                 =username.isDisplayed() && username.isEnabled();
		boolean actLPPassword                 =password.isDisplayed() && password.isEnabled();
		boolean actLPCompanyDDList            =companyDropDownList.isDisplayed() && companyDropDownList.isEnabled();
		boolean actLPSignInBtn                =signIn.isDisplayed() && signIn.isEnabled();
		boolean actRestoreBtn                 =loginPageRestoreBtn.isDisplayed();
		
		boolean expLPUserName                 =true;
		boolean expLPPassword                 =true;
		boolean expLPCompanyDDList            =true;
		boolean expLPSignInBtn                =true;
		boolean expRestroreBtn                =true;
		
		System.out.println("********************************** checkClickOnCancelButtonInRestoreCompany  *******************************************************");
		System.out.println("Login Page User Name Value Actual        :"+actLPUserName                 + "  Value Expected : " + expLPUserName);
		System.out.println("Login Page Password Value Actual         :"+actLPPassword                 + "  Value Expected : " + expLPPassword);
		System.out.println("Login Page Company DDList Value Actual   :"+actLPCompanyDDList            + "  Value Expected : " + expLPCompanyDDList);
		System.out.println("Login Page SignIn Btn Value Actual       :"+actLPSignInBtn                + "  Value Expected : " + expLPSignInBtn);
		System.out.println("Login Page Restore Btn Value Actual      :"+actRestoreBtn                 + "  Value Expected : " + expRestroreBtn);
		
		if(actLPUserName==expLPUserName && actLPPassword==expLPPassword && actLPCompanyDDList==expLPCompanyDDList && actLPSignInBtn==expLPSignInBtn)
	    {
		    excelReader.setCellData(xlfile, "Sheet1", 13, 9, resPass);
			return true;
        }
	    else 
	    {
			excelReader.setCellData(xlfile, "Sheet1", 13, 9, resFail);
			return false;
		
		 }
	 }
	 
	 
	 public boolean checkClickOnNoButtonInRestoreCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
		
		Thread.sleep(3000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginPageRestoreBtn));
		loginPageRestoreBtn.click();

        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(adminName));
		adminPassword.sendKeys("admin");
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginButton));
		loginButton.click();
		
		Thread.sleep(5000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreBrowseBtn));
		restoreBrowseBtn.click();
		
		Thread.sleep(5000);
		
		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\backupSelectionInRestoreCompany.exe");
		
		Thread.sleep(5000);

        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtnOfLoginPage));
		restoreCompanyBtnOfLoginPage.click();
		
	    Thread.sleep(5000);
		
		getWebDriverWait().until(ExpectedConditions.visibilityOf(restoreCompanyMessageNo));
	    restoreCompanyMessageNo.click();
	    
		getWaitForAlert();
				
		String actdata=getAlert().getText();

        // Restore company code : 030
		
		String actAppendCompanyCode="0"+actCompanyCodeIncreament+"0";
		
		System.out.println("actAppendCompanyCode"+actAppendCompanyCode);
		
		String expdata="Restore company code : "+ actAppendCompanyCode;
		
		System.out.println("expdata"+expdata);
		System.out.println("actdata"+actdata);
				
		if(actdata.equalsIgnoreCase(expdata))
		{
			
			getAlert().accept();
			System.out.println("Pass");
			return true;
		}
		else
		{
			getAlert().accept();
			System.out.println("Fail");
			return false;
			
		}
		
		
	 }
	 
	 
	 public boolean checkCompanyRestoreCompanyCodeIsDisplayInCompanyList() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
		
		Thread.sleep(5000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyDropDownList));
		companyDropDownList.click();
		
		Thread.sleep(2000);
		
		companyDropDownList.sendKeys(Keys.END);
		
		Thread.sleep(2000);
		
		companyDropDownList.sendKeys(Keys.ENTER);
		
		Thread.sleep(2000);
		
		String actCompanyName=companyDropDownList.getText();
		
		System.out.println("actCompanyName"+actCompanyName);
	
		String[] lines = actCompanyName.split("\n");

		System.out.println("First Company");
	
		String firstLine=lines[0];
        
        String actFirstCompanyCode = "";
        
        String firstCompCode=firstLine.substring(firstLine.length()-4);
        
        actFirstCompanyCode = firstCompCode.substring(1, 2);
        
        System.out.println("actFirstCompanyCode"+actFirstCompanyCode);
		
        System.out.println("Last Company");
			
        String lastLine=lines[lines.length-2];
        
        String actCompanyCode = "";
        
        CompCode=lastLine.substring(lastLine.length()-4);
        
        System.out.println("CompCode"+CompCode);
        
        actCompanyCode = CompCode.substring(1, 2);
         
        System.out.println("actCompanyCode"+actCompanyCode);
        
        String actCompanyCodeAfterRestore=lastLine;
        String expCompanyCodeAfterRestore="                            Automation Company ["+CompCode;
       
        
       
        System.out.println("********************************** checkCompanyRestoreCompanyCodeIsDisplayInCompanyList ********************************************");
		System.out.println("Admin Screen User Name Availablility Option Value Actual     :" + actCompanyCodeAfterRestore   +"  Value Expected :" + expCompanyCodeAfterRestore);
	
	    if(actCompanyCodeAfterRestore.equalsIgnoreCase(expCompanyCodeAfterRestore))
	    {
		    excelReader.setCellData(xlfile, "Sheet1", 13, 9, resPass);
			return true;
        }
	    else 
	    {
			excelReader.setCellData(xlfile, "Sheet1", 13, 9, resFail);
			return false;
		
		 }
	 }
	 
	 
	 @FindBy(xpath="//*[@id='id_focus_msgbox_main']/div[2]/button")
	 private static WebElement restoreErrorMessageCloseBtn;

	 
	 public boolean checkClickOnYesButtonInRestoreCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
		
		Thread.sleep(3000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginPageRestoreBtn));
		loginPageRestoreBtn.click();

        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(adminName));
		adminPassword.sendKeys("admin");
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginButton));
		loginButton.click();
		
		Thread.sleep(5000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreBrowseBtn));
		restoreBrowseBtn.click();
		
		Thread.sleep(5000);
		
		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\backupSelectionInRestoreCompany.exe");
		
		Thread.sleep(5000);

        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtnOfLoginPage));
		restoreCompanyBtnOfLoginPage.click();
		
	    Thread.sleep(5000);
			    
		getWebDriverWait().until(ExpectedConditions.visibilityOf(restoreCompanyMessageYes));
	    restoreCompanyMessageYes.click();
	    
		getWaitForAlert();
				
		String actdata=getAlert().getText();

        // Restore company code : 030
		
		actAppendCompanyCode="0"+actFirstCompanyCodeIncreament+"0";
		
		System.out.println("actAppendCompanyCode"+actAppendCompanyCode);
		
		String expdata="Restore company code : "+ actAppendCompanyCode;
		
		System.out.println("expdata"+expdata);
		System.out.println("actdata"+actdata);
				
		if(actdata.equalsIgnoreCase(expdata))
		{
			
			getAlert().accept();
			System.out.println("Pass");
			return true;
		}
		else
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreErrorMessageCloseBtn));
			restoreErrorMessageCloseBtn.click();
			System.out.println("Fail");
			return false;
		}
	 }
	 
	 
	 @FindBy(xpath="//span[@class='hidden-xs']")
	 private static WebElement userNameDisplay;

	 @FindBy(xpath="//*[@id='companyLogo']")
	 private static WebElement companyLogo;

	 @FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
	 private static WebElement companyName;

	 //Dash board Page
	 @FindBy(xpath="//*[@id='dashName']")
	 private static WebElement labelDashboard ;
			
	 @FindBy(xpath="//*[@id='Select_dash']")
	 private static WebElement selectDashboard ;
				
	 @FindBy(xpath="//*[@id='Dashboard_AddDash']")
	 private static WebElement newAddDashBoard;
			 
	 @FindBy(xpath="//*[@id='Dashboard_Dash_Config']")
	 private static WebElement dashboardCustomizationSettings;

	 @FindBy(xpath="//input[@id='donotshow']")
	 private static WebElement doNotShowCheckbox;
	    
	 @FindBy(xpath="//span[@class='pull-right']")
	 private static WebElement closeBtnInDemoPopupScreen;
	 
	 
	 public static boolean checkSignInBtToCheckRestoreYesWithValidUserNameAndValidPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			
			LoginPage lp=new LoginPage(getDriver());
			
			Thread.sleep(5000);
			
	        String unamelt=excelReader.getCellData("Sheet1", 9, 6);
	      
	        String pawslt=excelReader.getCellData("Sheet1", 10, 6);
	      
	        LoginPage.enterUserName(unamelt);
	     
	        LoginPage.enterPassword(pawslt);
	  
	        LoginPage.clickOnSignInBtn();  
	        
	        Thread.sleep(9000);
	       
	        //Login to company then message for demo is display
	        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(doNotShowCheckbox));
			doNotShowCheckbox.click();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnInDemoPopupScreen));
			closeBtnInDemoPopupScreen.click();
			
			Thread.sleep(2000);	
					        	
	    	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
	    	
	        Thread.sleep(2000);	
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo=userNameDisplay.getText();
			String expUserInfo="SU";
					   				    	
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
					    	
			String actGetCompCode=CompCode;
			String actCompanyCode1="";
			actCompanyCode1=CompCode.substring(0,3);
			
			System.out.println("actCompanyCode"+actCompanyCode1);
			
			System.out.println("actAppendCompanyCode value from other method"+actAppendCompanyCode);
			
			companyLogo.click();
			
			Thread.sleep(5000);
			
			String actGetLoginCompanyNameInformation=companyName.getText();
			//System.out.println("actGetLoginCompanyNameInformation"+actGetLoginCompanyNameInformation);
			String actGetLoginCompanyName=actGetLoginCompanyNameInformation.substring(0, 24);
			//System.out.println("actGetLoginCompanyName"+actGetLoginCompanyName);
			String expGetLoginCompanyName="Automation Company "+"("+actAppendCompanyCode+")";

			companyLogo.click();
					    	
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
					    	
			String actDashboardLabel=labelDashboard.getText();
		    String expDashboardLabel="Dashboard";

			System.out.println("***************************************************** checkSignInBtToCheckRestoreYesWithValidUserNameAndValidPassword *********************************************");
			System.out.println("Login To Company User Info Value Actual       :"+actUserInfo                       +"  Value Expected :" + expUserInfo);
			System.out.println("Login To Company Company Name Value Actual    :"+actGetLoginCompanyNameInformation +"  Value Expected :" + expGetLoginCompanyName);
			System.out.println("Login To Company Dashboard Label Value Actual :"+actDashboardLabel                 +"  Value Expected :" + expDashboardLabel);
			
			if(actUserInfo.equalsIgnoreCase(expUserInfo) && actGetLoginCompanyName.equalsIgnoreCase(expGetLoginCompanyName) && actDashboardLabel.equalsIgnoreCase(expDashboardLabel))
			{	
			    excelReader.setCellData(xlfile, "Sheet1", 9, 9, resPass);
			    return true;	 
			 }
			 else
			 {  			
				excelReader.setCellData(xlfile, "Sheet1", 10, 9, resFail);
				return false;
			 }
		}
		
	 
	  	
	 
		//Home Menu
		@FindBy (xpath="//*[@id='1']/div/span")
	    public static WebElement homeMenu;

		//Utilities Menu		
		@FindBy (xpath="//*[@id='20']/span")
	    private static WebElement utilities;

		@FindBy (xpath="//*[@id='106']/span")
	    private static WebElement dataStatistics;
	 	
		@FindBy(xpath="//*[@id='dataStatisticsBody']/tr")
		private static java.util.List<WebElement> rowsCountInDataStatistics;
		
		@FindBy(xpath="//div[contains(text(),'Refresh')]")
		public static WebElement refreshBtnInDataStatistics;
		 
		@FindBy(xpath="//div[contains(text(),'Backtrack')]")
		public static WebElement backtrackBtnInDataStatistics;
		 
		@FindBy(xpath="//div[contains(text(),'Filter')]")
		public static WebElement filterBtnInDataStatistics;
		 
		@FindBy(xpath="//div[contains(text(),'Sorting')]")
		public static WebElement sortingBtnInDataStatistics;
		 
		@FindBy(xpath="//div[contains(text(),'Close')]")
		public static WebElement closeBtnInDataStatistics;
		 
		@FindBy(xpath="//input[@id='txtSearchDSSReport']")
		public static WebElement searchOptionInDataStatistics;
		 
		
	 public static boolean checkDatatStatisticsInCompanyBackupRestoreWithYesOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
	
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
		utilities.click();
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataStatistics));
		dataStatistics.click();
		
		Thread.sleep(3000);
                
		int actRowSizeOfDataStatictics=rowsCountInDataStatistics.size();
		boolean actRefreshBtnInDataStatistics=refreshBtnInDataStatistics.isDisplayed();
		boolean actBacktrackBtnInDataStatistics=backtrackBtnInDataStatistics.isDisplayed();
		boolean actFilterBtnInDataStatistics=filterBtnInDataStatistics.isDisplayed();
		boolean actSortingBtnInDataStatistics=sortingBtnInDataStatistics.isDisplayed();
		boolean actCloseBtnInDataStatistics=closeBtnInDataStatistics.isDisplayed();
		boolean actSearchBtnInDataStatistics=searchOptionInDataStatistics.isDisplayed();
		
		int expRowSizeOdDataStatictics=0;
		boolean expRefreshBtnInDataStatistics=true;
		boolean expBacktrackBtnInDataStatistics=true;
		boolean expFilterBtnInDataStatistics=true;
		boolean expSortingBtnInDataStatistics=true;
		boolean expCloseBtnInDataStatistics=true;
		boolean expSearchBtnInDataStatistics=true;
		
	
        System.out.println("****************************************************** checkDatatStatistics ********************************************************************");
		System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actRowSizeOfDataStatictics  +"  Value Expected :"+ expRowSizeOdDataStatictics);
		System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actRefreshBtnInDataStatistics  +"  Value Expected :"+ expRefreshBtnInDataStatistics);
		System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actBacktrackBtnInDataStatistics  +"  Value Expected :"+ expBacktrackBtnInDataStatistics);
		System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actFilterBtnInDataStatistics  +"  Value Expected :"+ expFilterBtnInDataStatistics);
		System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actSortingBtnInDataStatistics  +"  Value Expected :"+ expSortingBtnInDataStatistics);
		System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actCloseBtnInDataStatistics  +"  Value Expected :"+ expCloseBtnInDataStatistics);
		System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actSearchBtnInDataStatistics  +"  Value Expected :"+ expSearchBtnInDataStatistics);
		
		
    
	    if(actRowSizeOfDataStatictics==expRowSizeOdDataStatictics)
	    {
		    excelReader.setCellData(xlfile, "Sheet1", 13, 9, resPass);
			return true;
        }
	    else 
	    {
			excelReader.setCellData(xlfile, "Sheet1", 13, 9, resFail);
			return false;
		
		 }
	 }
	 
	 
	 
	 public static boolean checkCloseBtnInDataStatistics() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		

        Thread.sleep(2000);

        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnInDataStatistics));
        closeBtnInDataStatistics.click();
        
        
        
        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));

		boolean actlabelDashboard=labelDashboard.isDisplayed();
		boolean actselectDashboard=selectDashboard.isDisplayed();
		boolean actnewAddDashBoard=newAddDashBoard.isDisplayed();
		boolean actdashboardCustomizationSettings=dashboardCustomizationSettings.isDisplayed();
		
		boolean explabelDashboard=true;
		boolean expselectDashboard=true;
		boolean expnewAddDashBoard=true;
		boolean expdashboardCustomizationSettings=true;
		
		System.out.println("labelDashboard  : "+actlabelDashboard+" "+explabelDashboard);
		System.out.println("selectDashboard  : "+actselectDashboard+" "+expselectDashboard);
		System.out.println("newAddDashBoard  : "+actnewAddDashBoard+" "+expnewAddDashBoard);
		System.out.println("dashboardCustomizationSettings  : "+actdashboardCustomizationSettings+" "+expdashboardCustomizationSettings);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
		if(labelDashboard.isDisplayed() && selectDashboard.isDisplayed() && newAddDashBoard.isDisplayed() && dashboardCustomizationSettings.isDisplayed())
		{
			System.out.println("test pass : Document History Pop Displayed");
			excelReader.setCellData(xlfile, "Sheet6", 83, 9, resPass);
			return true;
		}
		else
		{
			System.out.println("test Fail : Document History Pop NOT Displayed");
			excelReader.setCellData(xlfile, "Sheet6", 83, 9, resFail);
			return false;
		}
	 }
	 
	 
	 
	 @FindBy(xpath="//*[@id='userprofile']/li/span[2]")
	 private static WebElement logoutOption;
			
		
		//Logout and Login as Authorize User
		public boolean checkSignOutLoginCompanyWithYesOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

			  getWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
			  userNameDisplay.click();

			  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			  logoutOption.click();
			 
				boolean verifyFiledsStatus=true;
				
				getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
				
				
				boolean actusername=username.isDisplayed();
				boolean actpassword=password.isDisplayed();
				boolean actcompanyDropDownList=companyDropDownList.isDisplayed();
				boolean actsignIn=signIn.isDisplayed();
				
				// Expected
				boolean expusername=true;
				boolean exppassword=true;
				boolean expcompanyDropDownList=true;
				boolean expsignIn=true;
				
				System.out.println("************************************* checkSignOut  *********************************");
				System.out.println("username  : "+actusername+" Value Expected : "+expusername);
				System.out.println("password  : "+actpassword+" Value Expected : "+exppassword);
				System.out.println("companyDropDownList  : "+actcompanyDropDownList+" Value Expected : "+expcompanyDropDownList);
				System.out.println("signIn  : "+actsignIn+" Value Expected : "+expsignIn);
				
				if(actusername==expusername && actpassword==exppassword && actcompanyDropDownList==expcompanyDropDownList
						&& actsignIn==expsignIn)
				{

					excelReader.setCellData(xlfile, "Sheet1", 167, 9, resPass);
					return true;
				}
				else
				{
					excelReader.setCellData(xlfile, "Sheet1", 167, 9, resFail);
					return false;
			    }
		}
		
		
		
		
		 public static boolean checkSignInBtToCheckRestoreNoWithValidUserNameAndValidPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		 {
				
				excelReader=new ExcelReader(POJOUtility.getExcelPath());
				xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
				
				
				LoginPage lp=new LoginPage(getDriver());
				
				Thread.sleep(5000);
				
				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyDropDownList));
				companyDropDownList.click();
				Thread.sleep(2000);
				companyDropDownList.sendKeys(Keys.END);
				Thread.sleep(2000);
				companyDropDownList.sendKeys(Keys.ENTER);
				Thread.sleep(2000);
				
				
				
		        String unamelt=excelReader.getCellData("Sheet1", 9, 6);
		      
		        String pawslt=excelReader.getCellData("Sheet1", 10, 6);
		      
		        LoginPage.enterUserName(unamelt);
		     
		        LoginPage.enterPassword(pawslt);
		  
		        
		       
				LoginPage.clickOnSignInBtn();  
		        
		        Thread.sleep(9000);
		       
		        //Login to company then message for demo is display
		        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(doNotShowCheckbox));
				doNotShowCheckbox.click();
				
				Thread.sleep(2000);
				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnInDemoPopupScreen));
				closeBtnInDemoPopupScreen.click();
				
				Thread.sleep(2000);	
						        	
		    	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		    	
		        Thread.sleep(2000);	
				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
				String actUserInfo=userNameDisplay.getText();
				String expUserInfo="SU";
						   				    	
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
						    	
				String actGetCompCode=CompCode;
				String actCompanyCurrentCode="";
				actCompanyCurrentCode=CompCode.substring(0,3);
				
				System.out.println("actCompanyCode"+actCompanyCurrentCode);
				
				System.out.println("actAppendCompanyCode value from other method"+actAppendCompanyCode);
				
				companyLogo.click();
				
				Thread.sleep(5000);
				
				String actGetLoginCompanyNameInformation=companyName.getText();
				//System.out.println("actGetLoginCompanyNameInformation"+actGetLoginCompanyNameInformation);
				String actGetLoginCompanyName=actGetLoginCompanyNameInformation.substring(0, 24);
				//System.out.println("actGetLoginCompanyName"+actGetLoginCompanyName);
				String expGetLoginCompanyName="Automation Company "+"("+actCompanyCurrentCode+")";

				companyLogo.click();
						    	
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
						    	
				String actDashboardLabel=labelDashboard.getText();
			    String expDashboardLabel="Dashboard";

				System.out.println("***************************************************** checkSignInBtToCheckRestoreNoWithValidUserNameAndValidPassword *********************************************");
				System.out.println("Login To Company User Info Value Actual       :"+actUserInfo                       +"  Value Expected :" + expUserInfo);
				System.out.println("Login To Company Company Name Value Actual    :"+actGetLoginCompanyNameInformation +"  Value Expected :" + expGetLoginCompanyName);
				System.out.println("Login To Company Dashboard Label Value Actual :"+actDashboardLabel                 +"  Value Expected :" + expDashboardLabel);
				
				if(actUserInfo.equalsIgnoreCase(expUserInfo) && actGetLoginCompanyName.equalsIgnoreCase(expGetLoginCompanyName) && actDashboardLabel.equalsIgnoreCase(expDashboardLabel))
				{	
				    excelReader.setCellData(xlfile, "Sheet1", 9, 9, resPass);
				    return true;	 
				 }
				 else
				 {  			
					excelReader.setCellData(xlfile, "Sheet1", 10, 9, resFail);
					return false;
				 }
			}

		
		 
		 public static boolean checkDatatStatisticsInCompanyBackupRestoreWithNoOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		 {
			 
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			

			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();
		
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
			utilities.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataStatistics));
			dataStatistics.click();
			
			Thread.sleep(3000);
	                
			int actRowSizeOfDataStatictics=rowsCountInDataStatistics.size();
			boolean actRefreshBtnInDataStatistics=refreshBtnInDataStatistics.isDisplayed();
			boolean actBacktrackBtnInDataStatistics=backtrackBtnInDataStatistics.isDisplayed();
			boolean actFilterBtnInDataStatistics=filterBtnInDataStatistics.isDisplayed();
			boolean actSortingBtnInDataStatistics=sortingBtnInDataStatistics.isDisplayed();
			boolean actCloseBtnInDataStatistics=closeBtnInDataStatistics.isDisplayed();
			boolean actSearchBtnInDataStatistics=searchOptionInDataStatistics.isDisplayed();
			
			int expRowSizeOdDataStatictics=0;
			boolean expRefreshBtnInDataStatistics=true;
			boolean expBacktrackBtnInDataStatistics=true;
			boolean expFilterBtnInDataStatistics=true;
			boolean expSortingBtnInDataStatistics=true;
			boolean expCloseBtnInDataStatistics=true;
			boolean expSearchBtnInDataStatistics=true;
			
		
	        System.out.println("****************************************************** checkDatatStatistics ********************************************************************");
			System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actRowSizeOfDataStatictics  +"  Value Expected :"+ expRowSizeOdDataStatictics);
			System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actRefreshBtnInDataStatistics  +"  Value Expected :"+ expRefreshBtnInDataStatistics);
			System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actBacktrackBtnInDataStatistics  +"  Value Expected :"+ expBacktrackBtnInDataStatistics);
			System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actFilterBtnInDataStatistics  +"  Value Expected :"+ expFilterBtnInDataStatistics);
			System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actSortingBtnInDataStatistics  +"  Value Expected :"+ expSortingBtnInDataStatistics);
			System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actCloseBtnInDataStatistics  +"  Value Expected :"+ expCloseBtnInDataStatistics);
			System.out.println("Open Data Staticstics Of Overide Company Value Actual    :"+ actSearchBtnInDataStatistics  +"  Value Expected :"+ expSearchBtnInDataStatistics);
			
			
	    
		    if(actRowSizeOfDataStatictics==expRowSizeOdDataStatictics)
		    {
			    excelReader.setCellData(xlfile, "Sheet1", 13, 9, resPass);
				return true;
	        }
		    else 
		    {
				excelReader.setCellData(xlfile, "Sheet1", 13, 9, resFail);
				return false;
			
			 }
		 }
		 
		 
		 
		 public static boolean checkCloseBtnInDataStatisticsCompanyWithNoOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		 {
			 
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile="E:\\FocusSmoke\\FocusAI\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
	        Thread.sleep(2000);

	        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnInDataStatistics));
	        closeBtnInDataStatistics.click();
	        
	        
	        
	        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));

			boolean actlabelDashboard=labelDashboard.isDisplayed();
			boolean actselectDashboard=selectDashboard.isDisplayed();
			boolean actnewAddDashBoard=newAddDashBoard.isDisplayed();
			boolean actdashboardCustomizationSettings=dashboardCustomizationSettings.isDisplayed();
			
			boolean explabelDashboard=true;
			boolean expselectDashboard=true;
			boolean expnewAddDashBoard=true;
			boolean expdashboardCustomizationSettings=true;
			
			System.out.println("labelDashboard  : "+actlabelDashboard+" "+explabelDashboard);
			System.out.println("selectDashboard  : "+actselectDashboard+" "+expselectDashboard);
			System.out.println("newAddDashBoard  : "+actnewAddDashBoard+" "+expnewAddDashBoard);
			System.out.println("dashboardCustomizationSettings  : "+actdashboardCustomizationSettings+" "+expdashboardCustomizationSettings);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
			if(labelDashboard.isDisplayed() && selectDashboard.isDisplayed() && newAddDashBoard.isDisplayed() && dashboardCustomizationSettings.isDisplayed())
			{
				System.out.println("test pass : Document History Pop Displayed");
				excelReader.setCellData(xlfile, "Sheet6", 83, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("test Fail : Document History Pop NOT Displayed");
				excelReader.setCellData(xlfile, "Sheet6", 83, 9, resFail);
				return false;
			}
		 }
		 
		 
		    //Logout and Login as Authorize User
			public boolean checkSignOutLoginCompanyWithNoOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				excelReader=new ExcelReader(POJOUtility.getExcelPath());
				xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

				  getWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
				  userNameDisplay.click();

				  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
				  logoutOption.click();
				 
					boolean verifyFiledsStatus=true;
					
					getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
					
					
					boolean actusername=username.isDisplayed();
					boolean actpassword=password.isDisplayed();
					boolean actcompanyDropDownList=companyDropDownList.isDisplayed();
					boolean actsignIn=signIn.isDisplayed();
					
					// Expected
					boolean expusername=true;
					boolean exppassword=true;
					boolean expcompanyDropDownList=true;
					boolean expsignIn=true;
					
					System.out.println("************************************* checkSignOut  *********************************");
					System.out.println("username  : "+actusername+" Value Expected : "+expusername);
					System.out.println("password  : "+actpassword+" Value Expected : "+exppassword);
					System.out.println("companyDropDownList  : "+actcompanyDropDownList+" Value Expected : "+expcompanyDropDownList);
					System.out.println("signIn  : "+actsignIn+" Value Expected : "+expsignIn);
					
					if(actusername==expusername && actpassword==exppassword && actcompanyDropDownList==expcompanyDropDownList
							&& actsignIn==expsignIn)
					{

						excelReader.setCellData(xlfile, "Sheet1", 167, 9, resPass);
						return true;
					}
					else
					{
						excelReader.setCellData(xlfile, "Sheet1", 167, 9, resFail);
						return false;
				    }
			}
				
			
			
			@FindBy(xpath="//*[@id='RestoreCompanyModal']/div/div/div[2]/button[1]")
			private static WebElement OverwriteYesBtn;
			
			@FindBy(xpath="//*[@id='RestoreCompanyModal']/div/div/div[2]/button[2]")
			private static WebElement OverwriteNoBtn;
			
			@FindBy(xpath="//*[@id='RestoreCompanyModal']/div/div/div[2]/button[3]")
			private static WebElement OverwriteCancelBtn;
			
			
			
			
			 public static boolean checkRestoreCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
			 {
	 			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
	 			userNameDisplay.click();
	 			
	 			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
	 			logoutOption.click();
	 			
	 			Thread.sleep(2000);
	 			 
					getWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginPageRestoreBtn));
					loginPageRestoreBtn.click();
					
					getWebDriverWait().until(ExpectedConditions.elementToBeClickable(adminName));
					adminPassword.sendKeys("focus");
					
					getWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginButton));
					loginButton.click();
					
					Thread.sleep(2000);
					
			        boolean actRestoreCompanyBtnOfLoginPage  =restoreCompanyBtnOfLoginPage.isDisplayed();
			        boolean actRestoreInputTxt               =restoreInputTxt.isDisplayed();
			        boolean actRestoreBrowseBtn              =restoreBrowseBtn.isDisplayed();
			        
			        boolean expRestoreCompanyBtnOfLoginPage  =true;
			        boolean expRestoreInputTxt               =true;
			        boolean expRestoreBrowseBtn              =true;
					
			        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreBrowseBtn));
					restoreBrowseBtn.click();
					
					Thread.sleep(3000);
				     
				     Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\MaintainQuantityandValuebyWarehouse.exe");
				     
				     Thread.sleep(6000);
				     
				     getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtnOfLoginPage));
					 restoreCompanyBtnOfLoginPage.click();
					 
					 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(OverwriteYesBtn));
					 OverwriteYesBtn.click();
				     
			    	 getWaitForAlert();
			    	 String actAlertTxt = getAlert().getText();
			    	 getAlert().accept();
			 		 
			 		 System.err.println(actAlertTxt);
			 		 
			 		String expAlertTxt = "Restore company code : 030";
				    	 
				    Thread.sleep(3000);
				    
					if(actAlertTxt.equalsIgnoreCase(expAlertTxt))
					{
						return true;
					}
					else
					{
						return false;
					}
			     }

	 		
	 		
			 @FindBy(xpath = "//ul[@id='navigation_menu']/li/a/div/div")
	 		 private static List<WebElement> menusList;
			 
			 @FindBy(xpath = "//*[@id='mainHeader_MainLayout']/nav/div/ul/li")
			 private static List<WebElement> navbarList;
	 				
	 		 @FindBy (xpath="//*[@id='idquickAccessMenus_Main']/span[1]")
	 		 private static WebElement aboutQuickAcessMenu;
	 		
	 		 @FindBy (xpath="//*[@id='id_ai_search_input']")
	 		 private static WebElement aiSearchTextBox;
	 		
	 		 @FindBy (xpath="//*[@id='mainHeader_MainLayout']/nav/div/ul/li[1]/table/tbody/tr/td[2]/button")
	 		 private static WebElement aiSearchGoButton;
	 		
	 		 @FindBy (xpath="//*[@id='txtSearchMenu_MainLayout']")
	 		 private static WebElement searchTxtBox;
	 	
	 		 @FindBy (xpath="//*[@id='mainHeader_MainLayout']/nav/div/ul/li[3]/a/i")
	 		 private static WebElement productTour;
	 	
	 		 @FindBy (xpath="//*[@id='mainHeader_MainLayout']/nav/div/ul/li[4]/a/i")
	 		 private static WebElement alertsOptions;
	 	
	 		 @FindBy (xpath="//*[@id='idRecentMenus_Main']/i")
	 		 private static WebElement recentMenus;
	 		
	 		@FindBy (xpath="//div[@id='ReIndexingError_Modal']/div/div")
	 		private static WebElement reindexingPopup;

	 		@FindBy (xpath="(//div[@id='ReIndexingError_Modal']//div/button)[1]")
	 		private static WebElement reindexingPopupCloseBtn;

	 		@FindBy (xpath="(//div[@id='ReIndexingError_Modal']//div/button)[2]")
	 		private static WebElement reindexingPopupCancelBtn;

	 		@FindBy (xpath="//div[@id='ReIndexingError_Modal']//div/input")
	 		private static WebElement reindexingPopupSaveBtn;
	 		
	 		
	 		
	 		public static boolean checkLoginAfterRestoringBatchNegativeOptionsBackup() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	 		{
	 			Thread.sleep(3000);
	 			
	 			getDriver().navigate().refresh();
	 			
	 			LoginPage lp=new LoginPage(getDriver()); 

	 			String unamelt="su";

	 			String pawslt="su";

	 			boolean qflag=true;

	 			lp.enterUserName(unamelt);

	 			lp.enterPassword(pawslt);

	 			Thread.sleep(2000);
	 			
	 			
	 			 Select company = new Select(companyDropDownList);
	 			 company.selectByValue("108");
		        
	 			 Thread.sleep(2000);
	 			
	 			
	 			lp.clickOnSignInBtn();

	 			Thread.sleep(15000);

	 			try 
	 			{
	 				if (reindexingPopup.isDisplayed()) 
	 				{
	 					System.out.println("Reindexing Log is Displaying in Restore Company");

	 					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reindexingPopupCancelBtn));
	 					reindexingPopupCancelBtn.click();

	 					Thread.sleep(2000);

	 					lp.clickOnSignInBtn();
	 				}
	 				else
	 				{
	 					System.out.println("Reindexing Log is Not Displaying in Restore Company");
	 				}

	 				Thread.sleep(2000);
	 				
	 				String actUserInfo1=userNameDisplay.getText();

	 				System.out.println("User Info  : "+actUserInfo1);

	 				System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

	 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	 				companyLogo.click();

	 				String getCompanyTxt1=companyName.getText();
	 				String getLoginCompanyName1=getCompanyTxt1.substring(0, 19);
	 				System.out.println("company name  :  "+ getLoginCompanyName1);
	 				companyLogo.click();

	 				String expUserInfo1           ="SU";
	 				String expLoginCompanyName1   ="Maintain Quantity a";

	 				System.out.println("UserInfo1             : "+actUserInfo1            +" Value Expected : "+expUserInfo1);
	 				System.out.println("LoginCompanyName1     : "+getLoginCompanyName1    +" Value Expected : "+expLoginCompanyName1);

	 				if(actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1))
	 				{
	 					qflag=true;
	 				}
	 				else
	 				{
	 					qflag=false;
	 				}
	 			} 
	 			catch (Exception e) 
	 			{
	 				System.err.println(e.getMessage());
	 				
	 				Thread.sleep(2000);
	 				
	 				String actUserInfo1=userNameDisplay.getText();

	 				System.out.println("User Info  : "+actUserInfo1);

	 				System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

	 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	 				companyLogo.click();

	 				String getCompanyTxt1=companyName.getText();
	 				String getLoginCompanyName1=getCompanyTxt1.substring(0, 19);
	 				System.out.println("company name  :  "+ getLoginCompanyName1);
	 				companyLogo.click();

	 				String expUserInfo1           ="SU";
	 				String expLoginCompanyName1   ="Maintain Quantity a";

	 				System.out.println("UserInfo1             : "+actUserInfo1            +" Value Expected : "+expUserInfo1);
	 				System.out.println("LoginCompanyName1     : "+getLoginCompanyName1    +" Value Expected : "+expLoginCompanyName1);
	 			
	 				if(actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1))
	 				{
	 					qflag=true;
	 				}
	 				else
	 				{
	 					qflag=false;
	 				}
	 			}
	 			return qflag;
	 		}




			
			
			
			
			
			
		
	public RestoreCompanyPage(WebDriver driver)
	{
	  PageFactory.initElements(driver, this);	
	}

	
}
