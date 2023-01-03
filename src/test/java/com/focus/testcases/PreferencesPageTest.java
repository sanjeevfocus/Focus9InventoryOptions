package com.focus.testcases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.PreferencesPage;
//import com.focus.Pages.TCSTransactionDetailsPage;
import com.focus.base.BaseEngine;

public class PreferencesPageTest extends BaseEngine {
	PreferencesPage pp;
	//TCSTransactionDetailsPage tp;

	/*
	 * @Test(priority = 10) public void checkAvailablityOfTDSTab() throws
	 * EncryptedDocumentException, InvalidFormatException, IOException,
	 * InterruptedException {
	 * 
	 * pp = new PreferencesPage(getDriver());
	 * Assert.assertEquals(pp.checkAvailablityOfTDSTab(), true); }
	 * 
	 * @Test(priority = 11) public void clickOnVocherTypeAndvalidateTheVochers()
	 * throws EncryptedDocumentException, InvalidFormatException, IOException,
	 * InterruptedException {
	 * 
	 * pp = new PreferencesPage(getDriver());
	 * Assert.assertEquals(pp.clickOnVocherTypeAndvalidateTheVochers(), true); }
	 * 
	 * @Test(priority = 12) public void enterTdsTypeDetails() throws
	 * EncryptedDocumentException, InvalidFormatException, IOException,
	 * InterruptedException {
	 * 
	 * pp = new PreferencesPage(getDriver());
	 * Assert.assertEquals(pp.enterTdsTypeDetails(), true); }
	 */
	
	/*
	@Test(priority = 10)
	public void checkAvailablityOfTCSTab()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkAvailablityOfTCSTab(), true);
	}

	@Test(priority = 11)
	public void clickOnVocherTypeAndvalidateTheVochersInTCS()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.clickOnVocherTypeAndvalidateTheVochersInTCS(), true);
	}

	@Test(priority = 12)
	public void enterTcsTypeDetails()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.enterTcsTypeDetails(), true);
	}

	@Test(priority = 13)
	public void checkValidateEnableTCSTypeInCustomerTCSAccount()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkValidateEnableTCSTypeInCustomerTCSAccount(), true);
	}
		*/
	@Test(priority = 14)
	public void checkSavingSalesInvoiceVocher1()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkSavingSalesInvoiceVocher1(), true);
	}

	@Test(priority = 15)
	public void checkValuesInSavedSalesVoucher1()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkValuesInSavedSalesVoucher1(), true);
	}

	@Test(priority = 16)
	public void clickOnVocherTypeAndvalidateTheVochersInTCSAfterSavingSalesInvoceVocher()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.clickOnVocherTypeAndvalidateTheVochersInTCSAfterSavingSalesInvoceVocher(), true);
	}

	@Test(priority = 17)
	public void checkSavingSalesInvoiceVocher2()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkSavingSalesInvoiceVocher2(), true);
	}

	@Test(priority = 18)
	public void checkValuesInSavedSalesVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkValuesInSavedSalesVoucher2(), true);
	}

	@Test(priority = 19)
	public void checkSavingSalesDebitNotes()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkSavingSalesDebitNotes(), true);
	}

	@Test(priority = 20)
	public void checkSavingReceipts1()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkSavingReceipts1(), true);
	}

	@Test(priority = 21)
	public void validateVocherInTCSJvAfterSavingReceipts1()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.validateVocherInTCSJvAfterSavingReceipts1(), true);
	}

	@Test(priority = 22)
	public void checkSavingReceipts2()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkSavingReceipts2(), true);
	}
/*
	@Test(priority = 23)
	public void checkTCSTransactionsDetailsForAllAvailbletDsSEctionTYpes()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		tp = new TCSTransactionDetailsPage(getDriver());
		Assert.assertEquals(tp.checkTCSTransactionsDetailsForAllAvailbleTcsSectionTypes(), true);
	}
	*/

	@Test(priority = 24)
	public void checkValidateExportingTCSTransactionDetails()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkValidateExportingTCSTransactionDetails(), true);
	}

	@Test(priority = 25)
	public void checkNValidateExportedTCSTransactionDetailsExcel()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkNValidateExportedTCSTransactionDetailsExcel(), true);
	}

	@Test(priority = 26)
	public void checkGenerateChallanForAllVouchersInTDStransactionDetails()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		pp = new PreferencesPage(getDriver());
		Assert.assertEquals(pp.checkGenerateChallanForAllVouchersInTDStransactionDetails(), true);
	}

}
