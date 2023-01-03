package com.focus.testcases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.RestoreCompanyPage;
import com.focus.base.BaseEngine;

public class RestoreCompanyTest extends BaseEngine
{
	RestoreCompanyPage RCP;
	
	
	@Test(priority=10)
	public void checkRestoreCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		RCP = new RestoreCompanyPage(getDriver());
		Assert.assertEquals(RCP.checkRestoreCompany(), true);
	}
	
	@Test(priority=11)
	public void checkLoginAfterRestoringBatchNegativeOptionsBackup() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		RCP = new RestoreCompanyPage(getDriver());
		Assert.assertEquals(RCP.checkLoginAfterRestoringBatchNegativeOptionsBackup(), true);
	}
	
	
	
}
