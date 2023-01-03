package com.focus.testcases;

import org.testng.Assert;
import org.testng.annotations.Test;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import com.focus.Pages.LoginPage;
import com.focus.base.BaseEngine;



public class LoginPageTest extends BaseEngine 
{
	
	private static LoginPage lp;

	  @Test(priority=3)
	  public void checkLoginPageAvailablityOfFields() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  
		  lp=new LoginPage(getDriver());
		  Assert.assertEquals(lp.checkLoginPageAvailablityOfFields(), true);
	  }
	       
	  @Test(priority=4)
	  public void checkCompanyTextDisplayInCompanyListDropdown() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  lp=new LoginPage(getDriver());
		  Assert.assertEquals(lp.checkCompanyTextDisplayInCompanyListDropdown(), true);  
	  }
	  
	  @Test(priority=5)
	  public void verifyTheNumberOfListInCompanyDropdownList() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  lp=new LoginPage(getDriver());
		  Assert.assertEquals(lp.verifyTheNumberOfListInCompanyDropdownList(), true);
	  }
	  
	  @Test(priority=6)
	  public void checkSignInBtnWithValidUserNameAndValidPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  lp=new LoginPage(getDriver());
		  Assert.assertEquals(lp.checkSignInBtnWithValidUserNameAndValidPassword(), true);
	  }
	  
	  
	  
	  
	  
		
		/*
		 * @Test(priority=6) public void checkUserNameInfoInHomePage() throws
		 * EncryptedDocumentException, InvalidFormatException, IOException,
		 * InterruptedException { lp=new LoginPage(getDriver());
		 * 
		 * Assert.assertEquals(lp.checkUserNameInfoInHomePage(), true); }
		 */
	  
	 
	   
	        
	  /*
	  @Test(priority=7)
	  public void checkClickOnLogoutOptionInUserNameInfo() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  lp=new LoginPage(getDriver());
		  Assert.assertEquals(LoginPage.checkClickOnLogoutOptionInUserNameInfo(), true);
	  }
	  */

	/*  //@Test(priority=3)
	  public int getTheCompanyListCountBeforeCompanyList()
	  {
		  int compCount=lp.verifyTheNumberOfListInCompanyDropdownList();
		  return compCount;
	  }
	  
	  
	  //@Test(priority=5)
	  public int getTheCompanyListCountAfterCompanyList()
	  {
		  int compCount=lp.verifyTheNumberOfListInCompanyDropdownList();
		  return compCount;
	  }
	  

	 //@Test(priority=26) 
	 public void verifyUserDropDownImage() throws EncryptedDocumentException, InvalidFormatException, IOException
	 {

		 Assert.assertEquals(LoginPage.verifyUserNameImage(), true);

	 }
	 
	 
	 //@Test(priority=27)
	 public void verifyLogoutInUserDropdownImage() throws EncryptedDocumentException, InvalidFormatException, IOException
	 {
		 
		 LoginPage.clickOnLogOut();
		 //Assert.assertEquals(LoginPage.verifyAvailablityOfFields(), true);
		 
	 }
	
	
	 
	 //@Test(priority=28)
	 public void verifySignInRememberMeOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 
		 lp=new LoginPage(getDriver());
         excelReader=new ExcelReader(POJOUtility.getExcelPath());
       
         String unamelt=excelReader.getCellData("Sheet1",1,0);
       
         String pawslt=excelReader.getCellData("Sheet1", 1, 1);
       
         System.out.println(unamelt);   
        
         lp.enterUserName(unamelt);
   
         LoginPage.enterPassword(pawslt);
        
         LoginPage.checkRememberMe();
    
         LoginPage.clickOnSignInBtn();
        
         Assert.assertEquals(LoginPage.verifyUserHomePageAndDashBoard(), true);
        
        
	 }

	 
	//@Test(priority=29)
	 public void verifyLogoutRememberMeOption() throws InterruptedException
	 {
		 LoginPage.clickOnLogoutWithOptions();
		 
		 //Assert.assertEquals(LoginPage.signInWithSuperUserBycheckingRememberMeAndLogout(), true);
		 
	 }
	

	
	//@Test(priority=30)
	public void verifySignInWithRememberMeOptionByUserLoginAgainAfterLogout() throws InterruptedException
	{
			
	 //  Assert.assertEquals(LoginPage.signInAgainWithSuperUserBycheckingRememberMeAfterLogout(), true);
	}
	
	//@Test(priority=31)
	public void verifySignInWithRememberMyPasswordOption() throws InterruptedException
	{
		//lp.signInWithRememberMyPassword();
		
		Assert.assertEquals(LoginPage.verifyUserHomePageAndDashBoard(), true);
	}

	 
	 //@Test(priority=32)
	 public void verifySignOutWithSignInWithRememberMyPassword() throws InterruptedException
	 {
		 LoginPage.clickOnLogoutWithOptions();
		 Assert.assertEquals(LoginPage.signOutWithSignInAsRememeberMyPassword(), true);
	 }
	 
	 
	 //@Test(priority=33)
	 public void verifyForgotPassword() throws InterruptedException
	 {
		 
	//	 Assert.assertEquals(lp.checkForgotMyPassword(), true);
	 }
	 */
	 
	 
	 
	 
}
