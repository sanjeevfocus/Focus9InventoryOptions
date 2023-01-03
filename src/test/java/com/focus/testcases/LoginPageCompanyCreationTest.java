package com.focus.testcases;

import org.testng.Assert;
import org.testng.annotations.Test;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import com.focus.supporters.ExcelReader;
import com.focus.Pages.LoginPage;
import com.focus.Pages.LoginPageCompanyCreationPage;
import com.focus.base.BaseEngine;



public class LoginPageCompanyCreationTest extends BaseEngine {


	
	private static LoginPage lp;
	private static LoginPageCompanyCreationPage lpcc;
	private static ExcelReader excelReader;
	private static String url;
	private static String lgnTitle;
	private static String getCompanyInfo;


	  @Test(priority=8)
	  public void checkCreateCompanyInLoginPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  //Saving the Count of Company List before Company Creation
		  LoginPage lp=new LoginPage(getDriver());
		  lp.verifyTheNumberOfListInCompanyDropdownList();
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkCreateCompanyInLoginPage(), true);
	  }
	  
	  
	  @Test(priority=9)
	  public void checkCompanyCreatePageFields() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkCompanyCreatePageFields(), true);
	  }
	  
	  @Test(priority=10)
	  public void verifyOkButtonCompanyNameBlakInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkOKButtonCompanyNameBlankInCreateCompany(), true);  
	  }
	   
	  @Test(priority=11)
	  public void verifyOkButtonPasswordBlakInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkOKButtonPasswordBlankInCreateCompany(), true);  
	  }
	  

	  @Test(priority=12)
	  public void verifyCalenderTypeCountAndDefaultValueInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkCalenderTypeCountWithDisplayDefaultValue(), true);  
	  }

	  @Test(priority=13)
	  public void verifyCalenderTypeValuesInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkCalenderTypeValues(), true);
		  
	  }
	  
	  
	  @Test(priority=14)
	  public void verifyAccountingDateFormatInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkAccountingDateFormat(), true);
		  
	  }
	  
	  
	  
	  @Test(priority=15)
	  public void verifyCountryIndiaDisplayInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkCountry(), true);
		  
	  }
	  
	  @Test(priority=16)
	  public void verifyCountryComboboxCountInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkCountOfCountryValues(), true);
		  
	  }
	  
	  
	  @Test(priority=17)
	  public void verifyGlobalCurrencyDefaltValueDisplayInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkGlobalCurrencyDefaultDisplayValue(), true);
		  
	  }
	  
	  @Test(priority=18)
	  public void verifyGlobalCurrencyListOfCurrencyDisplayInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkCountOfGlobalCurrencyValues(), true);
		  
	  }
	  
	  
	  
	  
	  @Test(priority=19)
	  public void verifyLanguageDefaultLanguageComboboxInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkLanguagesOptionDefaultEnglishIsCheck(), true);
		  
	  }
	  
	  
	  @Test(priority=20)
	  public void verifyDefaultLanguageComboboxBySelectLangaugesInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkDefaultLanguageBySelectingOtherLanguages(), true);
		  
	  }

	 
	  @Test(priority=21)
	  public void verifyDefaultLanguageComboboxByUnSelectLangaugesInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkDefaultLanguageByUnSelectingOtherLanguages(), true);
		  
	  }
	  
	  @Test(priority=22)
	  public void verifyCompanyLogoInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPageCompanyCreationPage lpcp=new LoginPageCompanyCreationPage(getDriver());
		  Assert.assertEquals(LoginPageCompanyCreationPage.checkClickOkButtonToDisplayCompanyCreationAlert(), true);
		  
	  }
	  
	  @Test(priority=23)
	  public void verifyLoginScreenAfterCompanyCreationInCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  LoginPage lp=new LoginPage(getDriver());
		  Assert.assertEquals(LoginPage.checkLoginScreenAfterCompanyCreation(),true);
		  
	  }
	  
	  @Test(priority=24)
	  public void verifyCompanyListAfterCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  LoginPage lp=new LoginPage(getDriver());
		  Assert.assertEquals(LoginPage.checkCompanyListAfterCompanyCreate(),true);

	  }
	  
	  @Test(priority=25)
	  public void verifySignInAfterCreateCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  LoginPage lp=new LoginPage(getDriver());
		  Assert.assertEquals(LoginPage.checkSignInForNewCompanyCreate(),true);

	  }
	  
	  
	  @Test(priority=26)
	  public void verifyLogoutAfterCreateCompanyAndSignIn() throws EncryptedDocumentException, InvalidFormatException, IOException
	  {
		  LoginPage lp=new LoginPage(getDriver());
		  Assert.assertEquals(LoginPage.checkLogoutAfterCompanyCreationSignIn(), true);

	  }
	  
	  
	  
	  
	  
}
