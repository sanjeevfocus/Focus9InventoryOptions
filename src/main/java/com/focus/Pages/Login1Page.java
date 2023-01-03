package com.focus.Pages;

import java.io.IOException;

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

public class Login1Page extends BaseEngine {
	
	public Login1Page(WebDriver driver)
	{
	  PageFactory.initElements(driver, this);	
	  
	 }
	
	@FindBy(xpath="//*[@id='txtUsername']")
	private static WebElement username;

	@FindBy(id="txtPassword")
	private static WebElement password;
	
    @FindBy(id="ddlCompany")
    private static WebElement companyDropDownList;

	@FindBy(id="btnSignin")
	private static WebElement signIn;
	
	
	
	public static boolean checkSignInBtnWithValidUserNameAndValidPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
    {
		
		
		System.out.println(
				"********************************  checkSignInBtnWithValidUserNameAndValidPassword    ********************************************");

		String unamelt = "su";

		String pawslt = "su";

		enterUserName(unamelt);

		enterPassword(pawslt);

		Select company = new Select(companyDropDownList);
		company.selectByValue("108");
		Thread.sleep(3000);

		HomePage page=clickOnSignInBtn();
		
		
		
		
		
		return false;

		
		
		
		
		
		
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
	
	
	public static HomePage clickOnSignInBtn()
	{
		System.out.println("********************************* clickOnSignInBtn Method Is Executing...................  ***********************************");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		
		signIn.click();
		return null;	
	}
	
}
