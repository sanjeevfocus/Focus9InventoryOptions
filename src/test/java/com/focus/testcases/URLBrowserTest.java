package com.focus.testcases;

import org.testng.Assert;
import org.testng.annotations.Test;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import com.focus.supporters.ExcelReader;
import com.focus.Pages.LoginPage;
import com.focus.base.BaseEngine;



public class URLBrowserTest extends BaseEngine {
	
	private static LoginPage lp;
	private static ExcelReader excelReader;
	private static String url;
	private static String lgnTitle;
	private static String getCompanyInfo;

	
	  @Test(priority=1)
	  public void verifyBrowserOpen() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
        Assert.assertEquals(LoginPage.checkBrowserIsOpen(), true);
		
	  }

	  @Test(priority=2)
      public void checkLoginPageTitleByURLInputInBrowser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
      {   
	    Assert.assertEquals(LoginPage.checkLoginPageTitleByURLInputInBrowser(), true);
	       
      }
	  
	  
	  
	  
}
	