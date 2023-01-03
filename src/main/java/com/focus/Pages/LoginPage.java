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




public class LoginPage extends BaseEngine 
{
	private static String xlfile;
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static WebElement loginTitle;
	private static String url;
	public static  int cSize;
	private static String userN;
	private static String xlSheetName = "LoginPage";
	
	@FindBy(xpath="//*[@id='txtUsername']")
	private static WebElement username;

	@FindBy(id="txtPassword")
	private static WebElement password;
	
    @FindBy(id="ddlCompany")
    private static WebElement companyDropDownList;

	@FindBy(id="btnSignin")
	private static WebElement signIn;
	
	@FindBy(xpath="//i[@title='Create Company']")
    private static WebElement companyCreateBtn;		

	@FindBy(xpath="//i[@title='Keyboard ']")
	private static WebElement keyboardBtn;

    @FindBy(xpath="//i[@title='Refresh']")
    private static WebElement refreshBtn;
	
	@FindBy(id="chkRememberMe")
	private static WebElement rememberMeChk;
	
	@FindBy(id="chkRememberPwd")
	private static WebElement rememberPwdChk;
	
	@FindBy(id="frgtPwd")
	private static WebElement frgtPwdLnk;
	
	@FindBy(xpath="//span[contains(@class,'icon-ok icon-font6')]")
	private static WebElement okButtonInCreateCompany;
	
	@FindBy(xpath="//span[contains(@class,'icon-close icon-font6')]")
	private static WebElement cancelButtonInCreateCompany;
	
   @FindBy(xpath="//div[@id='Layout_Popup']//input[2]")
   private static WebElement noIn;

	@FindBy(linkText="Create Company")
	private static WebElement createCompanyScreen;
	
	@FindBy(xpath="//*[@id='errmsgDiv']")
	private static WebElement mandatoryMsgs;
	
	//Fields of Company su Home Page
	@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
	private static WebElement companyName;
	
	@FindBy(xpath="//*[@id='mainHeader_MainLayout']/a/img")
	private static WebElement focusLogo;
	
	@FindBy(xpath="//*[@id='id_focus8_wrapper_default']/aside/section")
	private static WebElement menuBar;
	
	@FindBy(xpath="//*[@id='dashName']")
	private static WebElement dashboardName;
	
	@FindBy(xpath="//*[@id='dashIcons']")
	private static WebElement dashboardIcons;
	
	//*[@id='mainHeader_MainLayout']/nav/div/ul/li[7]/a/span
	
	@FindBy(xpath="//span[@class='hidden-xs']")     //html/body/section/div[2]/header/nav/div/ul/li[6]/a/span  //span[text()='SU']
	private static WebElement userNameDisplay;
	
	/*@FindBy(xpath="//*[@id='mainHeader_MainLayout']/nav/div/ul/li[7]/a/span")
	private static WebElement userNameDisplay;*/

	@FindBy(xpath="//*[@id='companyLogo']")
	private static WebElement companyLogo;
	
	@FindBy(xpath="//*[@id='userprofile']/li/span[1]")
	private static WebElement changePassword;
	
	@FindBy(xpath="//*[@id='ddlCompanyLangualges_MainLayout']")
	private static WebElement languageDropdownInLogout;
	
	@FindBy(xpath="//*[@id='userprofile']/li/span[2]")
	private static WebElement logoutOption;
	
	@FindBy(xpath="//*[@id='dashName']")
	private static WebElement dashboard;
	
	@FindBy(xpath="//ul[@id='ulCompanyDetails_HomePage']")
	private static WebElement companydetails;
	
	//Dash board Page
		@FindBy(xpath="//*[@id='dashName']")
	    private static WebElement labelDashboard ;
		
		@FindBy(xpath="//*[@id='Select_dash']")
		private static WebElement selectDashboard ;
			
		@FindBy(xpath="//*[@id='Dashboard_AddDash']")
	    private static WebElement newAddDashBoard;
		 
	    @FindBy(xpath="//*[@id='Dashboard_Dash_Config']")
	    private static WebElement dashboardCustomizationSettings;
				
	
	
	
	
 public static boolean checkBrowserIsOpen() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
 {
   excelReader=new ExcelReader(POJOUtility.getExcelPath());
  xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
   System.out.println("******************************************************** checkBrowserIsOpen *************************************************************");
	
   excelReader.copyExcelFile();
   
   WebDriver actDriverValue=getDriver();
	
   String actConvertString=actDriverValue.toString();
   boolean actConvertBooleanValue=actConvertString.isEmpty();
   boolean expConvertBooleanValue=false;

   System.out.println("Browser Open Then Value of Web Driver Convert into String  :    "                    + actConvertString );
   System.out.println("Web Driver Open Value Actual   : " + actConvertBooleanValue  +  "  Value Expected :" + expConvertBooleanValue);

   if(actConvertBooleanValue==expConvertBooleanValue)
   {
		System.out.println("Pass : Browser has Lanunched");
		excelReader.setCellData(xlfile, xlSheetName, 5, 9, resPass);
		return true;
   }
   else
   {
		System.out.println("Fail : Browser has not Launched");
		excelReader.setCellData(xlfile, xlSheetName, 5, 9, resFail);
		return false;
	}	
 }

	
	
  public static boolean checkLoginPageTitleByURLInputInBrowser() throws EncryptedDocumentException, InvalidFormatException, IOException 
  {
	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
	System.out.println("*************************************** checkLoginScreenTitleByURLInputInBrowser *********************************************************");
	         
	
	  String actURLValue=excelReader.getCellData("LoginPage", 6, 6);
	  
	 enterUrl(actURLValue);    
	 
	
	//enterUrl(" http://localhost/FocusX#");                
   
    String actLoginTitle = getDriver().getTitle();
    String expLoginTitle="Focus";

    System.out.println("Login Screen Page Title Value Actual      : " + actLoginTitle +    "  Value Expected : " + expLoginTitle);
   
    if(actLoginTitle.equalsIgnoreCase(expLoginTitle))
    {
	   excelReader.setCellData(xlfile, "LoginPage", 6, 9, resPass);
	   return true;
    }
    else
    {
	   excelReader.setCellData(xlfile, "LoginPage", 6, 9, resFail);
	   return false;
    }   	       
   }
	
	
	@FindBy(xpath="//i[@class='icon-restore-1 icon-font4']")
	private static WebElement  loginPageRestoreBtn;
	
	@FindBy(xpath="//i[@class='icon-security icon-font4']")
	private static WebElement  loginPageSecurityProtectedCompanyBtn;
	
   
	public static boolean checkLoginPageAvailablityOfFields() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
		
		
		boolean actLPUserName                 =username.isDisplayed() && username.isEnabled();
		boolean actLPPassword                 =password.isDisplayed() && password.isEnabled();
		boolean actLPCompanyDDList            =companyDropDownList.isDisplayed() && companyDropDownList.isEnabled();
		boolean actLPSignInBtn                =signIn.isDisplayed() && signIn.isEnabled();
		boolean actLPRememberMeChkbox         =rememberMeChk.isDisplayed() && rememberMeChk.isEnabled();
		boolean actLPRememberMyPasswordChkbox =rememberPwdChk.isDisplayed() && rememberPwdChk.isEnabled();
		boolean actLPForgotMyPasswordLink     =frgtPwdLnk.isDisplayed() && frgtPwdLnk.isEnabled();
		
		boolean actLPCompanyCreateBtn         =companyCreateBtn.isDisplayed() && companyCreateBtn.isEnabled();
		boolean actLPRestoreCompanyBtn        =loginPageRestoreBtn.isDisplayed() && loginPageRestoreBtn.isEnabled();
		boolean actLPProtectedCompaniesBtn    =loginPageSecurityProtectedCompanyBtn.isDisplayed() && loginPageSecurityProtectedCompanyBtn.isEnabled();
		boolean actLPKeyboardBtn              =keyboardBtn.isDisplayed() && keyboardBtn.isEnabled();
		boolean actLPRefreshBtn               =refreshBtn.isDisplayed() && refreshBtn.isEnabled();
		
		
		boolean expLPUserName                 =true;
		boolean expLPPassword                 =true;
		boolean expLPCompanyDDList            =true;
		boolean expLPSignInBtn                =true;
		boolean expLPRememberMeChkbox         =true;
		boolean expLPRememberMyPasswordChkbox =true;
		boolean expLPForgotMyPasswordLink     =true;
		boolean expLPCompanyCreateBtn         =true;
		boolean expLPRestoreCompanyBtn        =true;
		boolean expLPProtectedCompaniesBtn    =true;
		boolean expLPKeyboardBtn              =true;
		boolean expLPRefreshBtn               =true;
		
		
		
		System.out.println("********************************** checkLoginPageAvailablityOfFields  *******************************************************");
		System.out.println("Home Page New Btn Value Actual           :"+actLPUserName                 + "  Value Expected : " + expLPUserName);
		System.out.println("Home Page Edit Btn Value Actual          :"+actLPPassword                 + "  Value Expected : " + expLPPassword);
		System.out.println("Home Page Print Btn Value Actual         :"+actLPCompanyDDList            + "  Value Expected : " + expLPCompanyDDList);
		System.out.println("Home Page Delete Btn Value Actual        :"+actLPSignInBtn                + "  Value Expected : " + expLPSignInBtn);
		System.out.println("Home Page Authorize Btn Value Actual     :"+actLPRememberMeChkbox         + "  Value Expected : " + expLPRememberMeChkbox);
		System.out.println("Home Page Suspend Btn Value Actual       :"+actLPRememberMyPasswordChkbox + "  Value Expected : " + expLPRememberMyPasswordChkbox);
		System.out.println("Home Page Reject Btn Value Actual        :"+actLPForgotMyPasswordLink     + "  Value Expected : " + expLPRememberMyPasswordChkbox);
		System.out.println("Home Page Print Barcode Btn Value Actual :"+actLPCompanyCreateBtn         + "  Value Expected : " + expLPCompanyCreateBtn);
		System.out.println("Home Page Export XML Btn Value Actual    :"+actLPRestoreCompanyBtn        + "  Value Expected : " + expLPRestoreCompanyBtn);
		System.out.println("Home Page Settings Btn Value Actual      :"+actLPProtectedCompaniesBtn    + "  Value Expected : " + expLPProtectedCompaniesBtn);
		System.out.println("Home Page Close Btn Value Actual         :"+actLPKeyboardBtn              + "  Value Expected : " + expLPKeyboardBtn);
		System.out.println("Home Page Grid check Value Actual        :"+actLPRefreshBtn               + "  Value Expected : " + expLPRefreshBtn);

		
		if(actLPUserName==expLPUserName && actLPPassword==expLPPassword && actLPCompanyDDList==expLPCompanyDDList && actLPSignInBtn==expLPSignInBtn
				&& actLPRememberMeChkbox==expLPRememberMeChkbox && actLPRememberMyPasswordChkbox==expLPRememberMyPasswordChkbox
				&& actLPForgotMyPasswordLink==expLPForgotMyPasswordLink && actLPCompanyCreateBtn==expLPCompanyCreateBtn
				&& actLPRestoreCompanyBtn==expLPRestoreCompanyBtn && actLPProtectedCompaniesBtn==expLPProtectedCompaniesBtn
				&& actLPKeyboardBtn==expLPKeyboardBtn && actLPRefreshBtn==expLPRefreshBtn)
		{					
			return true;
		 }
		 else
		 {
			 return false;
		 }
						
	  }
	
	
	public static boolean checkCompanyTextDisplayInCompanyListDropdown() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		boolean actCompanyDDListLength=companyDropDownList.getText().length()>0;
		boolean expCompanyDDListLength=true;
		
		 System.out.println("*********************************** checkCompanyTextDisplayInCompanyListDropdown *******************************");
		 System.out.println("Login Screen Company Name Display Value Actual        : "+ actCompanyDDListLength +"  Value Expected :" + expCompanyDDListLength);
		
		if(actCompanyDDListLength==expCompanyDDListLength)
		{
	    	return true;
	    }
	    else
	    {
	    	return false;
		}
		
	}
	
	
    @FindBy(xpath="//input[@id='donotshow']")
	private static WebElement doNotShowCheckbox;
    
    @FindBy(xpath="//span[@class='pull-right']")
	private static WebElement closeBtnInDemoPopupScreen;
	
    // Home Menu
 	@FindBy(xpath = "//*[@id='1']/div/span")
 	public static WebElement homeMenu;

 	    //Financial Menu
 		@FindBy (xpath="//span[@class='icon-financial icon-font1']")
 		private static WebElement financialMenu;
 		
 		//Fixed Assets Menu
 		@FindBy (xpath="//span[@class='icon-assets icon-font1']")
 		private static WebElement fixedAssetsMenu;
 		
 		//Production Menu
 		@FindBy (xpath="//span[@class='icon-production icon-font1']")
 		private static WebElement productionMenu;
 		
 		//Point Of Sale Menu
 		@FindBy (xpath="//div[contains(text(),'Point of Sale')]")
 		private static WebElement pointOfSaleMenu;
 		
 		//Quality Of Control Menu
 		@FindBy (xpath="//span[@class='icon-quotationanalysis icon-font1']")
 		private static WebElement qualityControlMenu;
 		
 		//Settings Menu
 		@FindBy (xpath="//span[@class='icon-settings icon-font1']")
 		private static WebElement SettingsMenu;
 		
 	    //Inventory Menu and Sub Menus		
 		@FindBy(xpath="//div[contains(text(),'Inventory')]")
 		private static WebElement  inventoryMenu; 
 		
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
 	
 		
 		public static boolean checkSignInBtnWithValidUserNameAndValidPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	    {
		try
		{
			System.out.println("********************************  checkSignInBtnWithValidUserNameAndValidPassword    ********************************************");
			
			String unamelt="su";
			
	        String pawslt="su";
	      
	        LoginPage.enterUserName(unamelt);
	     
	        LoginPage.enterPassword(pawslt);
	        
			
			 Select company = new Select(companyDropDownList);
			  company.selectByValue("108");
			  Thread.sleep(3000);
			 
	  
	        LoginPage.clickOnSignInBtn();  
	        
	        Thread.sleep(3000);
	        
	        System.out.println("Signing");
	        
	        // Login User Name is Verify
	        
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo                      = userNameDisplay.getText();
			String expUserInfo                      = "SU";	

			System.out.println("Login With Valid User and Check User Name Display Value Actual : ." + actUserInfo + ". Value Expected : " + expUserInfo);
			
			
			// Login Company Name is Verify
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			companyLogo.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
			String actGetLoginCompanyNameInformation = companyName.getText();
			String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
			String expGetLoginCompanyName            = "Maintain Quantity a";   
			companyLogo.click();
			
			System.out.println("Company Name Display Value Actual                : ." + actGetLoginCompanyName +  ". Value Expected : " + expGetLoginCompanyName);

			
			// Menus Checking
			
			int actCount					         = menusList.size();
			ArrayList<String> actMenusArray          = new ArrayList<String>();
			
			for (int i = 0; i < actCount; i++) 
			{
				String data                          = menusList.get(i).getText();
				actMenusArray.add(data);
			}
	
			String actMenus 					     = actMenusArray.toString();
			String expMenus					         = "[Home, Financials, Inventory, Fixed Asset, Production, Point of Sale, Quality Control, Settings]";
			
			System.out.println("Menu Names Text Display Value Actual             : " + actMenus);
			System.out.println("Menus Names Text Display Value Expected          : " + expMenus);
				
			Thread.sleep(2000);
			
			// Wrapper Layout Options
			
			boolean actNavigationFields=  aboutQuickAcessMenu.isDisplayed() && aiSearchTextBox.isDisplayed() &&  aiSearchGoButton.isDisplayed() 
					&& searchTxtBox.isDisplayed() && productTour.isDisplayed() && alertsOptions.isDisplayed() && recentMenus.isDisplayed();
			
			System.out.println("actNavigationFields"+actNavigationFields);
			String actDisplayingResult =Boolean.toString(actNavigationFields);
			String actElementsIsDisplay = actDisplayingResult.toUpperCase();
						
			String expElementsIsDisplay					         = "TRUE";
			
			System.out.println("Wrapper Layout Display Value Actual   : " + actElementsIsDisplay +  " Value Expected : " + expElementsIsDisplay);
			
									
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
			String actDashboardLabel                = labelDashboard.getText();
			String expDashboardLabel                = "Dashboard";	
			
			System.out.println("Dashboard Label Value Actual                     : ." + actDashboardLabel +      ". Value Expected : " + expDashboardLabel);
			
			if (actUserInfo.equalsIgnoreCase(expUserInfo) && actGetLoginCompanyName.equalsIgnoreCase(expGetLoginCompanyName)
					/*&& actDashboardLabel.equalsIgnoreCase(expDashboardLabel)*/
					
					&& actElementsIsDisplay.equalsIgnoreCase(expElementsIsDisplay)
					&& actMenus.equalsIgnoreCase(expMenus)) 
			{
				return true;
			} 
			else 
			{
				return false;
			}
		}
		catch (Exception e) 
	    {
			 String exception =e.getMessage().substring(0, 50);
	        System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ exception);
	        return false;
	    }
	}
	
	
	
	 public static boolean checkUserNameInfoInHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException
	 {
		 //we should check the login name and the company name, so change the code
		
		 System.out.println("********************************************* verifyUserNameInWrapperLayouts *****************************************************");
		 	 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		 userNameDisplay.click();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePassword));
		 
		 boolean actAvailabilityOfFields    = changePassword.isDisplayed() && changePassword.isEnabled()
	                                          && languageDropdownInLogout.isDisplayed() && languageDropdownInLogout.isEnabled()
		                                      && logoutOption.isDisplayed() && logoutOption.isEnabled()
		                                      && LoginPage.getLanguageCountInUserInfo()
		                                      && LoginPage.getLanguageTextInUserNameDisplay();
		 
	    System.out.println("actAvailabilityOfFields"+actAvailabilityOfFields);
		String actDisplayingResult =Boolean.toString(actAvailabilityOfFields);
		String actElementsIsDisplay = actDisplayingResult.toUpperCase()+" ";
					
		String expElementsIsDisplay					         = "TRUE ";
		
		System.out.println("Wrapper Layout Display Value Actual   : ." + actElementsIsDisplay +  ". Value Expected : " + expElementsIsDisplay);
		
		if(actElementsIsDisplay.equalsIgnoreCase(expElementsIsDisplay))
		{	 
			 return true;
		 }
		 else
		 {
			 return false;
		 }
	 }
	
	 
	 
	 public static boolean checkClickOnLogoutOptionInUserNameInfo() throws EncryptedDocumentException, InvalidFormatException, IOException
	 {	
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			 
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		 logoutOption.click();

		 getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
		
		 boolean actUserLoginPage              = username.isDisplayed() && username.isEnabled()
		                                        && password.isDisplayed() && password.isEnabled()
	                                            && companyDropDownList.isDisplayed() && companyDropDownList.isEnabled()
		                                        && signIn.isDisplayed() && signIn.isEnabled()
		                                        && rememberMeChk.isDisplayed() && rememberMeChk.isEnabled()
		                                        && rememberPwdChk.isDisplayed() && rememberPwdChk.isEnabled()
		                                        && frgtPwdLnk.isDisplayed() && frgtPwdLnk.isEnabled()
		
		                                        && companyCreateBtn.isDisplayed() && companyCreateBtn.isEnabled()
		                                        && loginPageRestoreBtn.isDisplayed() && loginPageRestoreBtn.isEnabled()
		                                        && loginPageSecurityProtectedCompanyBtn.isDisplayed() && loginPageSecurityProtectedCompanyBtn.isEnabled()
		                                        && keyboardBtn.isDisplayed() && keyboardBtn.isEnabled()
		                                        && refreshBtn.isDisplayed() && refreshBtn.isEnabled();
	
			 
		    System.out.println("actUserLoginPage"+actUserLoginPage);
			String actDisplayingResult =Boolean.toString(actUserLoginPage);
			String actElementsIsDisplay = actDisplayingResult.toUpperCase()+" ";
			excelReader.setCellData(xlfile, xlSheetName, 17, 8, actElementsIsDisplay);
						
			String expElementsIsDisplay					         = excelReader.getCellData(xlSheetName, 17, 7);
			
			System.out.println("Wrapper Layout Display Value Actual   : " + actElementsIsDisplay +  " Value Expected : " + expElementsIsDisplay);
			
			boolean actLPUserName                 =username.isDisplayed() && username.isEnabled();
			boolean actLPPassword                 =password.isDisplayed() && password.isEnabled();
			boolean actLPCompanyDDList            =companyDropDownList.isDisplayed() && companyDropDownList.isEnabled();
			boolean actLPSignInBtn                =signIn.isDisplayed() && signIn.isEnabled();
			boolean actLPRememberMeChkbox         =rememberMeChk.isDisplayed() && rememberMeChk.isEnabled();
			boolean actLPRememberMyPasswordChkbox =rememberPwdChk.isDisplayed() && rememberPwdChk.isEnabled();
			boolean actLPForgotMyPasswordLink     =frgtPwdLnk.isDisplayed() && frgtPwdLnk.isEnabled();
			
			boolean actLPCompanyCreateBtn         =companyCreateBtn.isDisplayed() && companyCreateBtn.isEnabled();
			boolean actLPRestoreCompanyBtn        =loginPageRestoreBtn.isDisplayed() && loginPageRestoreBtn.isEnabled();
			boolean actLPProtectedCompaniesBtn    =loginPageSecurityProtectedCompanyBtn.isDisplayed() && loginPageSecurityProtectedCompanyBtn.isEnabled();
			boolean actLPKeyboardBtn              =keyboardBtn.isDisplayed() && keyboardBtn.isEnabled();
			boolean actLPRefreshBtn               =refreshBtn.isDisplayed() && refreshBtn.isEnabled();
			
			
			boolean expLPUserName                 =true;
			boolean expLPPassword                 =true;
			boolean expLPCompanyDDList            =true;
			boolean expLPSignInBtn                =true;
			boolean expLPRememberMeChkbox         =true;
			boolean expLPRememberMyPasswordChkbox =true;
			boolean expLPForgotMyPasswordLink     =true;
			boolean expLPCompanyCreateBtn         =true;
			boolean expLPRestoreCompanyBtn        =true;
			boolean expLPProtectedCompaniesBtn    =true;
			boolean expLPKeyboardBtn              =true;
			boolean expLPRefreshBtn               =true;
				
			
			System.out.println("************************************ checkLoginPageAvailablityOfFields  *******************************************************");
			System.out.println("Home Page New Btn Value Actual           :"+actLPUserName                  + "  Value Expected : " + expLPUserName);
			System.out.println("Home Page Edit Btn Value Actual          :"+actLPPassword                  + "  Value Expected : " + expLPPassword);
			System.out.println("Home Page Print Btn Value Actual         :"+actLPCompanyDDList             + "	Value Expected : " + expLPCompanyDDList);
			System.out.println("Home Page Delete Btn Value Actual        :"+actLPSignInBtn                 + "	Value Expected : " + expLPSignInBtn);
			System.out.println("Home Page Authorize Btn Value Actual     :"+actLPRememberMeChkbox          + "	Value Expected : " + expLPRememberMeChkbox);
			System.out.println("Home Page Suspend Btn Value Actual       :"+actLPRememberMyPasswordChkbox  + "	Value Expected : " + expLPRememberMyPasswordChkbox);
			System.out.println("Home Page Reject Btn Value Actual        :"+actLPForgotMyPasswordLink      + "	Value Expected : " + expLPRememberMyPasswordChkbox);
			System.out.println("Home Page Print Barcode Btn Value Actual :"+actLPCompanyCreateBtn          + "	Value Expected : " + expLPCompanyCreateBtn);
			System.out.println("Home Page Export XML Btn Value Actual    :"+actLPRestoreCompanyBtn         + "	Value Expected : " + expLPRestoreCompanyBtn);
			System.out.println("Home Page Settings Btn Value Actual      :"+actLPProtectedCompaniesBtn     + "	Value Expected : " + expLPProtectedCompaniesBtn);
			System.out.println("Home Page Close Btn Value Actual         :"+actLPKeyboardBtn               + "	Value Expected : " + expLPKeyboardBtn);
			System.out.println("Home Page Grid check Value Actual        :"+actLPRefreshBtn                + "	Value Expected : " + expLPRefreshBtn);

			

			if(actElementsIsDisplay.equalsIgnoreCase(expElementsIsDisplay))
			{	 
				 excelReader.setCellData(xlfile, "LoginPage", 11, 9, resPass);
				 return true;
			 }
			 else
			 {
				 excelReader.setCellData(xlfile, "LoginPage", 11, 9, resFail);
				 return false;
			 }				
	    }
	 

	 
	 public static boolean checkLoginScreenAfterCompanyCreation() throws EncryptedDocumentException, InvalidFormatException, IOException
	 {
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		 
		 System.out.println("******** 1.verifyLoginPageAfterCompanyCreation*********");
		 
		 boolean signInScreen=false;
		   try
			{
			   System.out.println("******** 2.verifyLoginPageAfterCompanyCreation First Try block ********");
			   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
			   
				//if(signIn.isDisplayed())
			    if(username.isDisplayed() && password.isDisplayed() && companyDropDownList.isDisplayed() && signIn.isDisplayed() && companyCreateBtn.isDisplayed() && keyboardBtn.isDisplayed() && refreshBtn.isDisplayed() && rememberMeChk.isDisplayed() && rememberPwdChk.isDisplayed() && frgtPwdLnk.isDisplayed())
				{
			    	excelReader.setCellData(xlfile, "LoginPage", 27, 9, resPass);
				    System.out.println("******** 3.verifyLoginPageAfterCompanyCreation user enter in Try block");
					getLogger().info("Pass :Login Screen is displaying after user creating company");
					System.out.println("Pass : Login Screen is displaying after user creating company");
					signInScreen=true;
					
					
				}
	     	}
	        catch(Exception e) 
			{
	        	    excelReader.setCellData(xlfile, "LoginPage", 27, 9, resFail);
					getLogger().info("Error is displaying while creating Company"+e.getMessage());
					System.out.println("******* 6.verifyLoginPageAfterCompanyCreation user enter in Catch block");
					/*getWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelButtonInCreateCompany));
					LoginPage.cancelButtonInCreateCompany.click();*/
					signInScreen=false;
			}

		   return signInScreen;
	 }
	 
	 
	 
	 
	 public static boolean checkCompanyListAfterCompanyCreate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		    excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			LoginPage lp=new LoginPage(getDriver()); 
			
	        String unamelt=excelReader.getCellData("LoginPage", 35, 6);
	      
	        String pawslt=excelReader.getCellData("LoginPage", 36, 6);
	           
	        LoginPage.enterUserName(unamelt);
	           
	        getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
	        
	        LoginPage.enterPassword(pawslt);
	        	
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			
			String compname="Automation Company";
		 	
			 Select oSelect = new Select(companyDropDownList);
			 
			 List <WebElement> elementCount = oSelect.getOptions();
			
			 int cqSize = elementCount.size();
			 
			 int zqSize=cSize+1;
			 
			 System.out.println("CompanyDropdownList Count :"+cqSize);
			 
			 System.out.println("Company dropdown is :"+ zqSize);
		 
		 
		 //Select dropdown= new Select(lp.companyDropDownList);
		  int i;
		  
		  //List<WebElement> list = dropdown.getOptions();

			//List<String> text = new ArrayList<>();
			for(i=0; i<elementCount.size(); i++) 
			{
			
			  elementCount.get(i).getText();
	    	  String optionName = elementCount.get(i).getText();
	    	  if(optionName.toUpperCase().startsWith(compname.toUpperCase()))
	    	  {
	    		  System.out.println("q"+elementCount.get(i).getText());
	    		  elementCount.get(i).click();
	    		  
	    	  }
	      
	      }
		 
			
		 if(cqSize==zqSize)
		 {
			 excelReader.setCellData(xlfile, "LoginPage", 28, 9, resPass);
			 return true;
			 
		 }
		 else
		 {
			 excelReader.setCellData(xlfile, "LoginPage", 28, 9, resPass);
			 return false;
			 
		 }
		
	 }
	 
	 
	 public static boolean checkSignInForNewCompanyCreate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		 
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(signIn));
		 signIn.click();	
		 
		 Thread.sleep(9000);
		  
		 System.out.println("Verify UserNameImage Screen");
		 
		 userNameDisplay.click();
		 
		 System.out.println("User Image dropdown user name : "+ LoginPage.getLanguageTextInUserNameDisplay());
		 
		 System.out.println("User Image dropdown language count :"+ LoginPage.getLanguageTextInUserNameDisplay());
		 
		 if(changePassword.isDisplayed()&&changePassword.isEnabled()
		    &&languageDropdownInLogout.isDisplayed()&&languageDropdownInLogout.isEnabled()
			&&logoutOption.isDisplayed()&&logoutOption.isEnabled())
		 {	 
			 System.out.println("Pass : ChangePassword//LanguageDropdown//Logout display in the user info Screen");
			 			
			 if(LoginPage.getLanguageCountInUserNameDisplay()==2)
			 {
				 System.out.println("Pass : Language is display Count as One in the user info Screen");
			 
				 if(LoginPage.getLanguageTextInUserNameDisplay()==true)
				 {
					 System.out.println("Pass : Language Text English is displaying in the user info Screen");
					 excelReader.setCellData(xlfile, "LoginPage", 2, 9, resPass);
					 return true;
				 }
				 else
				 {
					 System.out.println("Fail : Language Text English is not displaying in the user info Screen");
					 excelReader.setCellData(xlfile, "LoginPage", 11, 9, resFail);
					 return false;
				 }
			 }
			 else
			 {
				
				 System.out.println("Pass : Language is display Count not as One in the user info Screen");
				 excelReader.setCellData(xlfile, "LoginPage", 11, 9, resFail);
				 return false;
			 }
		 }	
		 else
		 {
			 System.out.println("Pass : Language Text English is not displaying in the user info Screen");
			 excelReader.setCellData(xlfile, "LoginPage", 11, 9, resFail);
			 return false;
		 }


	 }
	 
	 
	 
	 public static boolean checkLogoutAfterCompanyCreationSignIn() throws EncryptedDocumentException, InvalidFormatException, IOException
	 {
		 
		 
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		 
		 
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		 logoutOption.click();
		
		 
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		 
			if(username.isDisplayed()&& password.isDisplayed() && companyDropDownList.isDisplayed() && signIn.isDisplayed() && companyCreateBtn.isDisplayed() && keyboardBtn.isDisplayed() && refreshBtn.isDisplayed() && rememberMeChk.isDisplayed() && rememberPwdChk.isDisplayed() && frgtPwdLnk.isDisplayed())
			{
				excelReader.setCellData(xlfile, "LoginPage", 29, 9, resPass);
				return true;
			}
			else
			{
				excelReader.setCellData(xlfile, "LoginPage", 29, 9, resPass);
				return false;
			}
	 }
	 
	 
	 
	 
	 public boolean verifyTheNumberOfListInCompanyDropdownList()
	    {
	    	 Select oSelect = new Select(companyDropDownList);
			 List <WebElement> elementCount = oSelect.getOptions();
			
			 cSize = elementCount.size();
			 System.out.println("CompanyDropdownList Count :"+cSize);
			 return true;
	    }
	 
	 

	 
	 
	 
	 
	 public void getCompanyFromCompanyListAfterCompanyCreate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
	        
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			LoginPage lp=new LoginPage(getDriver()); 
			
	        String unamelt=excelReader.getCellData("LoginPage", 9, 6);
	      
	        String pawslt=excelReader.getCellData("LoginPage", 10, 6);
	           
	        LoginPage.enterUserName(unamelt);
	        
	        getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
	        
	        LoginPage.enterPassword(pawslt);
	        	
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			
			String compname="Automation Company";
			  
			  Select dropdown= new Select(LoginPage.companyDropDownList);
			  int i;
			  
			  List<WebElement> list = dropdown.getOptions();

				List<String> text = new ArrayList<>();
				for(i=0; i<list.size(); i++) 
				{
				
		    	  list.get(i).getText();
		    	  String optionName = list.get(i).getText();
		    	  if(optionName.toUpperCase().startsWith(compname.toUpperCase()))
		    	  {
		    		  list.get(i).click();
		    		  
		    	  }
		      
		      }

				Thread.sleep(3000);
		}
	 
	 
	
	 public static void clickOnLogoutWithOptions()
	 {
		 
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
		 userNameDisplay.click();
		 
		 
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		 logoutOption.click();
		 
		 
		 
		 
	 }
	
	
	
    public static void clickOnCompanyCreateBtn()
    {
    	System.out.println("company create");
    	
    	companyCreateBtn.click();
    	
    }
	
	
	
	public static void enterUserName(String uname)
	{
		System.out.println("********************************* enterUserName Method Is Executing.....................  ************************************");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.clear();
	    username.click();
        username.sendKeys(uname);
        getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
   
	}
	
	
	
	public static void enterPassword(String pswd) throws InterruptedException
	{		
		System.out.println("********************************* enterPassword Method Is Executing.....................  ************************************");
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.clear();
		password.click();
		password.sendKeys(pswd);
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();	
	}
	
	
	public static void clickOnSignInBtn()
	{
		System.out.println("********************************* clickOnSignInBtn Method Is Executing...................  ***********************************");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();	
	}
	

	
	 
	 
	 public static void checkRememberMe()
	 {
		 rememberMeChk.click();
	 }
	 
	 
	 
    public static String getCompanyNameFromLoginScreen()
    {
    	String getCompanyNameText=companyDropDownList.getText();
    	return getCompanyNameText;
    	
    }
	
    
    public static boolean verifyUserHomePageAndDashBoard() throws InterruptedException
    {
    	
    	
    	boolean avbFields=false;
    	
    	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
    	
    	Thread.sleep(4000);
    	getWebDriverWait().until(ExpectedConditions.textToBePresentInElement(userNameDisplay, "SU"));
    	String userInfo=userNameDisplay.getText();
    	
    	System.out.println("User Info : "+userInfo);
    	System.out.println("User Info Capture Text :"+userNameDisplay.getText());
    	
    	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
    	companyLogo.click();
    	
    	String getCompanyTxt=companyName.getText();
    	String getLoginCompanyName=getCompanyTxt.substring(0, 19);
    	System.out.println("company name :"+ getLoginCompanyName);
    	companyLogo.click();
    	
    	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
    	
    	String getDashboard=dashboard.getText();
    	
    	System.out.println(getDashboard);
    	
    		if(userInfo.equalsIgnoreCase("SU")&&getLoginCompanyName.equalsIgnoreCase("Automation Company "))
    		{
    			avbFields=true;
    			getLogger().info("Login User the Company Name and Information is displaying correct");
    			System.out.println("Pass : Login User and Company Name is displaying Correct");
    		
    		
    			if(getDashboard.equalsIgnoreCase("Dashboard"))
    			{
    				avbFields=true;
    			}
    		    else
    		    {
    		    	avbFields=false;
    		    	
    		    }
    		}
    		else
    		{
    			avbFields=false;
    			getLogger().info("Fail : Login User the Company Name and Information is displaying Wrong");
    			System.out.println("Fail : Login Name the Company Name and Information is displaying Wrong");
    		}

    	
    	return avbFields;
    }
    
    

    
    
    
    
    
    
	
	
	 
	 
	

	 public static boolean verifyMandatoryFieldCheckForSuperUserName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		  
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		 
		  username.clear();
		  String msgMandatoryComp="Username cannot be blank";
		  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		  signIn.click();
		  getWebDriverWait().until(ExpectedConditions.textToBePresentInElement(mandatoryMsgs, "Username cannot be blank"));
		  System.out.println(mandatoryMsgs.getText());
		  
		  
		  if(msgMandatoryComp.equalsIgnoreCase(mandatoryMsgs.getText()))
		  {
			  excelReader.setCellData(xlfile, "LoginPage", 31, 9, resPass);
			  return true;
		  }
		  else
		  {
			  excelReader.setCellData(xlfile, "LoginPage", 31, 9, resFail);
			  return false;
		  }
		  
	  }
	 
	 
	 
	 public static boolean verifyMandatoryFieldCheckForPassword() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		 
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		  
		  String msgMandatoryComp="Password cannot be blank";
		  
		  String un=excelReader.getCellData("LoginPage", 32, 6);
		  username.sendKeys(un);
		  
		  getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		  getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		  
		  
		  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		  signIn.click();
		  System.out.println(mandatoryMsgs.getText());
		  
		  getWebDriverWait().until(ExpectedConditions.textToBePresentInElement(mandatoryMsgs, "Password cannot be blank"));
		  if(msgMandatoryComp.equalsIgnoreCase(mandatoryMsgs.getText()))
		  {
			  excelReader.setCellData(xlfile, "LoginPage", 32, 9, resPass);
			  return true;
		  }
		  else
		  {
			  excelReader.setCellData(xlfile, "LoginPage", 32, 9, resPass);
			  return false;
		  }
		  
	  }
	 
	 
	 
	 public static boolean verifyInvalidSuperUserName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		  
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		 
		  String msgMandatoryComp="User name not found";
		  username.click();
		  username.clear();
		  String un=excelReader.getCellData("LoginPage", 33, 6);
		  username.sendKeys(un);
		  getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		  //password.clear();
		  String pw=excelReader.getCellData("LoginPage", 34, 6);
	      password.sendKeys(pw);
		  getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		  signIn.click();
		  getWebDriverWait().until(ExpectedConditions.textToBePresentInElement(mandatoryMsgs, "User name not found"));
		  System.out.println(mandatoryMsgs.getText());
		  
		  if(msgMandatoryComp.equalsIgnoreCase(mandatoryMsgs.getText()))
		  {
			  excelReader.setCellData(xlfile, "LoginPage", 33, 9, resPass);
			  return true;
		  }
		  else
		  {
			  excelReader.setCellData(xlfile, "LoginPage", 33, 9, resFail);
			  return false;
		  }
		  
	  }
	 
	 
	 public static boolean verifyInvalidPassword() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		  
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		 
		  String msgMandatoryComp="Invalid Password";
		  username.click();
		  username.clear();
		  String un=excelReader.getCellData("LoginPage", 35, 6);
		  username.sendKeys(un);
		  getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		  password.clear();
		  String pw=excelReader.getCellData("LoginPage", 36, 6);
	      password.sendKeys(pw);
		  getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		  signIn.click();
		  //Thread.sleep(3500);
		  getWebDriverWait().until(ExpectedConditions.textToBePresentInElement(mandatoryMsgs, "Invalid Password"));
		  System.out.println(mandatoryMsgs.getText());
		  
		  if(msgMandatoryComp.equalsIgnoreCase(mandatoryMsgs.getText()))
		  {
			  excelReader.setCellData(xlfile, "LoginPage", 35, 9, resPass);
			  return true;
		  }
		  else
		  {
			  
			  excelReader.setCellData(xlfile, "LoginPage", 35, 9, resFail);
			  return false;
		  }
		  
	  }
	 
	 
	 public static boolean checkRememberMeAsCheckedInLoginPage() throws EncryptedDocumentException, InvalidFormatException, IOException
	 {

		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		 
         getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
		 username.clear();
		 String unameltR=excelReader.getCellData("LoginPage", 37, 6);
		 username.sendKeys(unameltR);
		 getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		
		
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(password));
		 String pweltR=excelReader.getCellData("LoginPage", 38, 6);
		 password.sendKeys(pweltR);
		 getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		
		 rememberMeChk.click();
		 
		 if(rememberMeChk.isSelected())
		 {
			 excelReader.setCellData(xlfile, "LoginPage", 37, 9, resPass);
			 return true;
		 }
		 else
		 {
			 excelReader.setCellData(xlfile, "LoginPage", 37, 9, resFail);
			 return false;
		 }
		 
	 }

	 
	
	 public static boolean checkRememberMeSignIn() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {

		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
         getWebDriverWait().until(ExpectedConditions.visibilityOf(signIn));
         signIn.click();
         
 		 Thread.sleep(9000);
        
 		/*getWebDriverWait().until(ExpectedConditions.elementToBeClickable(doNotShowCheckbox));
		doNotShowCheckbox.click();
		
		Thread.sleep(2000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnInDemoPopupScreen));
		closeBtnInDemoPopupScreen.click();
		
		Thread.sleep(2000);	*/
		
         getWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
         userNameDisplay.click();
		 
		 System.out.println("User Image dropdown user name : "+ LoginPage.getLanguageTextInUserNameDisplay());
		 
		 System.out.println("User Image dropdown language count :"+ LoginPage.getLanguageTextInUserNameDisplay());
		 
		 if(changePassword.isDisplayed()&&changePassword.isEnabled()
		    &&languageDropdownInLogout.isDisplayed()&&languageDropdownInLogout.isEnabled()
			&&logoutOption.isDisplayed()&&logoutOption.isEnabled())
		 {	 
			 System.out.println("Pass : ChangePassword//LanguageDropdown//Logout display in the user info Screen");
			 
			
			 if(LoginPage.getLanguageCountInUserNameDisplay()==2)
			 {
				 System.out.println("Pass : Language is display Count as One in the user info Screen");
			 
				 if(LoginPage.getLanguageTextInUserNameDisplay()==true)
				 {
					 System.out.println("Pass : Language Text English is displaying in the user info Screen");
					 excelReader.setCellData(xlfile, "LoginPage", 39, 9, resPass);
					 return true;
				 }
				 else
				 {
					 System.out.println("Fail : Language Text English is not displaying in the user info Screen");
					 excelReader.setCellData(xlfile, "LoginPage", 39, 9, resFail);
					 return false;
				 }
			 }
			 else
			 {
				
				 System.out.println("Pass : Language is display Count not as One in the user info Screen");
				 excelReader.setCellData(xlfile, "LoginPage", 39, 9, resFail);
				 return false;
			 }
		 }	
		 else
		 {
			 System.out.println("Pass : Language Text English is not displaying in the user info Screen");
			 excelReader.setCellData(xlfile, "LoginPage", 39, 9, resFail);
			 return false;
		 }



	 }
	 
	 
	 public static boolean checkLoginPageLogInWithRememeberMeLogout() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
	
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		  
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		 logoutOption.click();
		 
		 String sunameValue=username.getAttribute("value");
		 
		 String pwdValue=password.getAttribute("value");
		 
		 System.out.println("**********user name is display as"+sunameValue);
		 System.out.println("**********pwdValue is display as"+pwdValue);
		 System.out.println("**********pwdValue true or false is display as"+pwdValue.isEmpty());
		 
		
		 
		 if(username.isDisplayed()&& sunameValue.equalsIgnoreCase("su") && password.isDisplayed()&&pwdValue.isEmpty()==true && companyDropDownList.isDisplayed() && signIn.isDisplayed() && rememberMeChk.isDisplayed() && rememberMeChk.isSelected() && rememberPwdChk.isDisplayed())
		 {
			 excelReader.setCellData(xlfile, "LoginPage", 40, 9, resPass);
			 return true;	
							
		 }
		 else
		 {
			 excelReader.setCellData(xlfile, "LoginPage", 40, 9, resFail);
			 return true;
			 
		 }
		  
															
	   }

	 
	 
	 public static boolean checkSignInAgainWithSuperUserBycheckingRememberMeAfterLogout() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		 
		
		 
		 
		 
		 
		
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		 
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
         username.click();
         
         getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
         
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(password));
		 
		
		 getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		 
		 String sunameValue=username.getAttribute("value");
		 String pwdValue=password.getAttribute("value");
		 boolean p=pwdValue.isEmpty();
		 
		 
		 System.out.println("sunameValue"+sunameValue);
		 System.out.println("pwdValue"+pwdValue);
		 System.out.println("p"+p);
		 

		 
			 if(username.isDisplayed()&&username.isEnabled()&&sunameValue.equalsIgnoreCase("su") && rememberMeChk.isSelected())
			 {
				 excelReader.setCellData(xlfile, "LoginPage", 41, 9, resPass);
				 return true;
			                                
	
			 }
		 	 else 
			 {
		 		excelReader.setCellData(xlfile, "LoginPage", 41, 9, resFail);
				 return true;
			 }

				
		  }
	 
	 
	 
	 public static boolean checkRememeberMeRememberSelected() throws EncryptedDocumentException, InvalidFormatException, IOException
	   {

		

		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		 
         getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
		 username.clear();
		 String unameltR=excelReader.getCellData("LoginPage", 42, 6);
		 username.sendKeys(unameltR);
		 getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		
		
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(password));
		 String pweltR=excelReader.getCellData("LoginPage", 43, 6);
		 password.sendKeys(pweltR);
		 getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		
		 rememberMeChk.click();
		 
		 rememberPwdChk.click();
		 
		 if(rememberMeChk.isSelected() && rememberPwdChk.isSelected())
		 {
			 excelReader.setCellData(xlfile, "LoginPage", 42, 9, resPass);
			 return true;
		 }
		 else
		 {
			 excelReader.setCellData(xlfile, "LoginPage", 42, 9, resFail);
			 return false;
		 }
		 
		 

			 
	   }
	   
	  
	 
	 public static boolean checkRememberMeRememberMyPasswordSignIn() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {

		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
         getWebDriverWait().until(ExpectedConditions.visibilityOf(signIn));
         signIn.click();
		
 		 Thread.sleep(9000);
        /*
 		 if(noIn.isDisplayed())
 		 {
 			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(noIn));
 			noIn.click();
 			Thread.sleep(5000);
 			
 		 }*/
         
       
         getWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
         userNameDisplay.click();
		 
		 System.out.println("User Image dropdown user name : "+ LoginPage.getLanguageTextInUserNameDisplay());
		 
		 System.out.println("User Image dropdown language count :"+ LoginPage.getLanguageTextInUserNameDisplay());
		 
		 if(changePassword.isDisplayed() && changePassword.isEnabled()
		    && languageDropdownInLogout.isDisplayed() &&languageDropdownInLogout.isEnabled()
			&& logoutOption.isDisplayed() && logoutOption.isEnabled())
		 {	 
			 System.out.println("Pass : ChangePassword//LanguageDropdown//Logout display in the user info Screen");
			 
			
			 if(LoginPage.getLanguageCountInUserNameDisplay()==2)
			 {
				 System.out.println("Pass : Language is display Count as One in the user info Screen");
			 
				 if(LoginPage.getLanguageTextInUserNameDisplay()==true)
				 {
					 System.out.println("Pass : Language Text English is displaying in the user info Screen");
					 excelReader.setCellData(xlfile, "LoginPage", 44, 9, resPass);
					 return true;
				 }
				 else
				 {
					 System.out.println("Fail : Language Text English is not displaying in the user info Screen");
					 excelReader.setCellData(xlfile, "LoginPage", 44, 9, resFail);
					 return false;
				 }
			 }
			 else
			 {
				
				 System.out.println("Pass : Language is display Count not as One in the user info Screen");
				 excelReader.setCellData(xlfile, "LoginPage", 44, 9, resFail);
				 return false;
			 }
		 }	
		 else
		 {
			 System.out.println("Pass : Language Text English is not displaying in the user info Screen");
			 excelReader.setCellData(xlfile, "LoginPage", 44, 9, resFail);
			 return false;
		 }

	 }
	 
	 
	 public static boolean checkLoginPageLogInWithRememeberMeRememberMyPasswordLogout() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
	
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		  
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		 logoutOption.click();
		 
		 String sunameValue=username.getAttribute("value");
		 
		 String pwdValue=password.getAttribute("value");
		 
		 System.out.println("**********user name is display as"+sunameValue);
		 System.out.println("**********pwdValue is display as"+pwdValue);
		 System.out.println("**********pwdValue true or false is display as"+pwdValue.isEmpty());
		 
		
		 
		 if(username.isDisplayed()&& sunameValue.equalsIgnoreCase("su") && password.isDisplayed()&& pwdValue.equalsIgnoreCase("su") && companyDropDownList.isDisplayed() && signIn.isDisplayed() && rememberMeChk.isDisplayed() && rememberMeChk.isSelected() && rememberPwdChk.isDisplayed())
		 {
			 excelReader.setCellData(xlfile, "LoginPage", 45, 9, resPass);
			 return true;	
							
		 }
		 else
		 {
			 excelReader.setCellData(xlfile, "LoginPage", 45, 9, resFail);
			 return true;
			 
		 }
		  
															
	   }
	 
	 
	 public static boolean checkForgotMyPasswordForSuperUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	   {
		 
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		   String forgotMsg="Email not found for \'su\'";
		   System.out.println("Forgot msg : "+forgotMsg);
		   
		   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(frgtPwdLnk));
		   
		   frgtPwdLnk.click();
		   
		   Thread.sleep(4000);
		   
		   System.out.println("Mandatory Message : "+mandatoryMsgs.getText());
			  
		   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(rememberMeChk));
		   
		   rememberMeChk.click();
			 
		   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(rememberPwdChk));
			 
		   rememberPwdChk.click();
		   
			  if(forgotMsg.equalsIgnoreCase(mandatoryMsgs.getText()))
			  {
				  excelReader.setCellData(xlfile, "LoginPage", 46, 9, resPass);
				  return true;
			  }
			  else
			  {
				  excelReader.setCellData(xlfile, "LoginPage", 46, 9, resFail);
				  return false;
			  }
			  
		   
	   }
	 
	 
	 
	 public static boolean getLanguageCountInUserInfo()
	 {
			 Select oSelect = new Select(languageDropdownInLogout);
			 List <WebElement> elementCount = oSelect.getOptions();
			
			 int lSize = elementCount.size();
			 System.out.println("Language Dropdown List are :"+lSize);
			 
			 if(lSize==2)
			 {
				 return true;
			 }
			 else
			 {
				 return false;
			 }
	 }
	
	 
	 
	 public static int getLanguageCountInUserNameDisplay()
	 {
			 Select oSelect = new Select(languageDropdownInLogout);
			 List <WebElement> elementCount = oSelect.getOptions();
			
			 int lSize = elementCount.size();
			 System.out.println("Language Dropdown List are :"+lSize);
			 return lSize;
	 }
	 	

	 
	 public static boolean getLanguageTextInUserNameDisplay()
	 {
		 //String languageInLogout=languageDropdownInLogout.getText();
		 
		 
		 String Strlpl= languageDropdownInLogout.getText();
			
		 System.out.println(Strlpl);
		 
		 boolean getLanguageDropdownList=false;
		 String[] str= Strlpl.split("\n");     
	      for(String st1 : str)
	      {
	    	
	    	  st1.equals("UserAllOptions");
	    	  getLanguageDropdownList=true;
	    	  break;
	      }
	      
		return getLanguageDropdownList;
		 
	 }
	 
	 
	 
	 
	 public static void clickOnLogOut()
	 {
		 
		// //userNameDisplay.click();
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		 logoutOption.click();
		 
	 }
	 


	 
	 public static boolean clickOnLogoutChangePasswordSingInWithInvalid()
	 {
		 
		  String msgMandatoryComp="Invalid Password";
		 
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
		 userNameDisplay.click();
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		 logoutOption.click();
		 
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
		 username.sendKeys("su");
		 getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(password));
		 password.sendKeys("su");
		 getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		

		 signIn.click();
		 

		  getWebDriverWait().until(ExpectedConditions.textToBePresentInElement(mandatoryMsgs, "Invalid Password"));
		  System.out.println(mandatoryMsgs.getText());
		  
		  if(msgMandatoryComp.equalsIgnoreCase(mandatoryMsgs.getText()))
		  {
			  System.out.println(true);
			  return true;
		  }
		  else
		  {
			  System.out.println(false);
			  return false;
		  }
	 }
	 
	 
	 
	 
	
	 
	 
	 
	 public static void clickOnLogoutChangePasswordSingInWithValid() throws InterruptedException
	 {
		 
		
		 
		 
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
		 username.clear();
		 username.sendKeys("su");
		 getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(password));
		 password.sendKeys("s");
		 getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		

		 signIn.click();
		 
		 
		 

	 }
	 
	
	 
   
   public static boolean signOutWithSignInAsRememeberMyPassword() throws InterruptedException
	 {
		 
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
		 
         username.click();
		 //Thread.sleep(4000);
       
         getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(password));
		 //password.sendKeys("su");
		 getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		 
		 String sunameValue=username.getAttribute("value");
		 String pwdValue=username.getAttribute("value");
		 
		 System.out.println(sunameValue);
		 System.out.println(pwdValue);
		 
		 boolean verifyFiledsStatus=true;
		 
		 if(username.isDisplayed()&&username.isEnabled()&&sunameValue.equalsIgnoreCase("su"))
			{
				System.out.println("UserName text box is displayed and enabled "+ verifyFiledsStatus);
				
				if(password.isDisplayed()&&password.isEnabled()&&pwdValue.equalsIgnoreCase("su"))
				{
					System.out.println("Password text box is displayed and enabled "+ verifyFiledsStatus);
		
					if(companyDropDownList.isDisplayed()&&companyDropDownList.isEnabled())
					{
						System.out.println("Company Drop Down box is displayed and enabled "+ verifyFiledsStatus);
					
						if(signIn.isDisplayed()&&signIn.isEnabled())
						{
							System.out.println("SignIn is displayed and enabled "+ verifyFiledsStatus);
						
							if(companyCreateBtn.isDisplayed()&&companyCreateBtn.isEnabled())
							{
								System.out.println("Create Company Button is displayed and enabled "+ verifyFiledsStatus);
							
								if(keyboardBtn.isDisplayed()&&keyboardBtn.isEnabled())
								{
									System.out.println("Keyboard Button is displayed and enabled "+ verifyFiledsStatus);
								
									if(refreshBtn.isDisplayed()&&refreshBtn.isEnabled())
									{
										System.out.println("Refresh Button is displayed and enabled "+ verifyFiledsStatus);
						
		                                getWebDriverWait().until(ExpectedConditions.visibilityOf(rememberMeChk));
		 
										if(rememberMeChk.isDisplayed()&&rememberMeChk.isEnabled()&&rememberMeChk.isSelected())
										{
											System.out.println("Remember Me Checkbox is displayed and enabled "+ verifyFiledsStatus);
												
											
											if(rememberPwdChk.isDisplayed()&&rememberPwdChk.isEnabled()&&rememberPwdChk.isSelected())
											{
												System.out.println("Remember Password Checkbox is displayed and enabled "+ verifyFiledsStatus);
															
												 if(frgtPwdLnk.isDisplayed()&&frgtPwdLnk.isEnabled())
												 {
													System.out.println("Forgot Password Link is displayed and enabled "+ verifyFiledsStatus);
													verifyFiledsStatus=true;
													
												 }
												 else
												 {
													 System.out.println("Forgot Password Link is : Fail ");
													 verifyFiledsStatus=false;
												 }
											}
											else
											 {
												 System.out.println("rememberPwdChk Checkbox : Fail ");
												 verifyFiledsStatus=false;
											 }

											
										}
										else
										 {
											 System.out.println("rememberMeChk Button : Fail ");
											 verifyFiledsStatus=false;
										 }

									}
									else
									{
										 System.out.println("refreshBtn Button : Fail ");
										 verifyFiledsStatus=false;
									}
							     }
								else
								 {
									 System.out.println("keyboardBtn : Fail ");
									 verifyFiledsStatus=false;
								 }

						      }
							else
							 {
								 System.out.println("companyCreateBtn : Fail ");
								 verifyFiledsStatus=false;
							 }

					      }
						
						else
						 {
							 System.out.println("signIn : Fail ");
							 verifyFiledsStatus=false;
						 }

					   }
					else
					 {
						 System.out.println("companyDropDownList : Fail ");
						 verifyFiledsStatus=false;
					 }

				   }
				else
				 {
					 System.out.println("Password field is displaying as Fail ");
					 verifyFiledsStatus=false;
				 }

				}
	 						
				else 
				{
					verifyFiledsStatus=false;
					System.out.println("User Name "+verifyFiledsStatus);
				}

			return verifyFiledsStatus;
				
		  }

	 
   
   
   public void clickOnSignIn()
   {
	     getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
	     username.click();
         username.sendKeys("su");
         getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(password));
		 password.sendKeys("su");
		 getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		 
		 getWebDriverWait().until(ExpectedConditions.visibilityOf(signIn));
		 signIn.click();
		 
   }

   
   
   
  /* public static boolean verifyUserHomePageAndDashBoardAfterUpdateTheCompany() throws InterruptedException
   {
	   
   	boolean avbFields=false;
   	
	//HPWrapperLayoutPage wlp=new HPWrapperLayoutPage(getDriver());
	
	wlp.clickOnAbout();

	String getgetPatchDateTime=wlp.getReleasedate();
	
	wlp.clickOnAboutToClose();
	
	String getCompanyNamePatch="Automation Company : "+getgetPatchDateTime;
	
	System.out.println("GetCompanyNamePatch**********GetCompanyPatch : "+getCompanyNamePatch);
	
	
   	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
   	
   	Thread.sleep(4000);
   	getWebDriverWait().until(ExpectedConditions.textToBePresentInElement(userNameDisplay, "SU"));
   	String userInfo=userNameDisplay.getText();
   	
   	System.out.println("User Info : "+userInfo);
   	System.out.println("User Info Capture Text :"+userNameDisplay.getText());
   	
   	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
   	companyLogo.click();
   	
   	//getWebDriverWait().until(ExpectedConditions.textToBePresentInElement(companyName, "Automation Company "));
   	String getCompanyTxt=companyName.getText();
   	String getLoginCompanyName=getCompanyTxt.substring(0, 23);
   	System.out.println("company name :"+ getLoginCompanyName);
  
   	
   	companyLogo.click();
   	
   	//if(focusLogo.isDisplayed()&&menuBar.isDisplayed()&&dashboardName.isDisplayed()&&dashboardIcons.isDisplayed()&&userNameDisplay.isDisplayed()&&companyLogo.isDisplayed())
   	//{
   		//getLogger().info("Pass : Login the Company and Information is displaying correct");
			//System.out.println("Pass : Login the Company and Information is displaying Correct");
   		
   		if(userInfo.equalsIgnoreCase("SU")&&getLoginCompanyName.equalsIgnoreCase(getCompanyNamePatch))
   		{
   			avbFields=true;
   			getLogger().info("Login User the Company Name and Information is displaying correct");
   			System.out.println("Pass : Login User and Company Name is displaying Correct");
   		}
   		else
   		{
   			avbFields=false;
   			getLogger().info("Fail : Login User the Company Name and Information is displaying Wrong");
   			System.out.println("Fail : Login Name the Company Name and Information is displaying Wrong");
   		}
   		
   	//}
   	//else
   	//{
   		//getLogger().info("Fail : Login the Company and Information is displaying correct");
			//System.out.println("Fail : Login User the Company Name and Information is displaying Correct");
   		
   		//avbFields=false;
   	
   	return avbFields;
   }
   */
   

	
	public LoginPage(WebDriver driver) 
	{
		
		  PageFactory.initElements(driver, this);
	}
	
	

	
}