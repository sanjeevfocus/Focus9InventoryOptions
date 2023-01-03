package com.focus.testcases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.ItemQueryPage;
import com.focus.base.BaseEngine;

public class ItemQueryPageTest extends BaseEngine {

	ItemQueryPage ip;
	
	@Test(priority = 12)
	public void validteItemQueryStocksInBatchItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException 
	{
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInBatchItem(), true);
	}
	
	@Test(priority = 13)
	public void validteItemQueryStocksInFifoItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException 
	{
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInFifoItem(), true);
	}
	
	@Test(priority = 14)
	public void validteItemQueryStocksInWAItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException 
	{
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInWAItem(), true);
	}
	
	@Test(priority = 15)
	public void checkStockLedgerReportForBrBatchItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForBrBatchItem(), true);
	}
	
	@Test(priority = 16)
	public void checkStockLedgerReportForFifoItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForFifoItem(), true);
	}
	
	@Test(priority = 17)
	public void checkStockLedgerReportForRmaItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForRmaItem(), true);
	}
	// Purchase voucher 
	
	@Test(priority = 18) // getting issue in  Rma item  Overall avg Rate not matching.
	public void checkSavePurchasesVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavePurchasesVoucher(), true);
	}
	
	@Test(priority = 19)
	public void validteItemQueryStocksInBatchItemAfterSavingPurchaseVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInBatchItemAfterSavingPurchaseVocher(), true);
	}

	@Test(priority = 20)
	public void validteItemQueryStocksInFifoItemAfterSavingPurchaseVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInFifoItemAfterSavingPurchaseVocher(), true);
	}

	@Test(priority = 21)
	public void validteItemQueryStocksInWAItemAfterSavingPurchaseVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInWAItemAfterSavingPurchaseVocher(), true);
	}
	
	@Test(priority = 22)
	public void checkStockLedgerReportFortBrBatchItemAfterSavingPurchaseVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportFortBrBatchItemAfterSavingPurchaseVocher(), true);
	}

	@Test(priority = 23)
	public void checkStockLedgerReportForFifoItemAfterSavingPurchaseVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForFifoItemAfterSavingPurchaseVocher(), true);
	}
		
	@Test(priority = 24)
	public void checkStockLedgerReportForRmaItemAfterSavingPurchaseVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForRmaItemAfterSavingPurchaseVocher(), true);
	}
	
	@Test(priority = 25)
	public void checkLoginToInvUserAndValidteItemQueryStocksInBatchItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkLoginToInvUserAndValidteItemQueryStocksInBatchItem(), true);
	}
	
	@Test(priority = 26)
	public void validteItemQueryStocksInFifoItemAfterLoginToInvUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInFifoItemAfterLoginToInvUser(), true);
	}

	@Test(priority = 27)
	public void validteItemQueryStocksInWAItemAfterLoginToInvUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInWAItemAfterLoginToInvUser(), true);
	}
	
	
	@Test(priority = 28)
	public void checkStockLedgerReportForBrBatchItemAfterLoginToInvUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForBrBatchItemAfterLoginToInvUser(), true);
	}

	@Test(priority = 29)
	public void checkStockLedgerReportForFifoItemAfterLoginToInvUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForFifoItemAfterLoginToInvUser(), true);
	}
	
	@Test(priority = 30)
	public void checkStockLedgerReportForRmaItemAfterLoginToInvUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForRmaItemAfterLoginToInvUser(), true);
	}
	
	// Sales invoice voucher         
	
	@Test(priority = 31)
	public void checkSavingSalesInvoiceVoucherBrBatchItemWithHydWarehouse()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingSalesInvoiceVoucherBrBatchItemWithHydWarehouse(), true);
	}
	
	@Test(priority = 32)
	public void checkSavingSalesInvoiceVoucherBrBatchItemWithSECWarehouse()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingSalesInvoiceVoucherBrBatchItemWithSECWarehouse(), true);
	}
	
	@Test(priority = 33)
	public void checkSavingSalesInvoiceVoucherFifoItemWith2Warehouse()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingSalesInvoiceVoucherFifoItemWith2Warehouse(), true);
	}
	
	@Test(priority = 34)  // getting issue in  Rma item  Overall avg Rate not matching.
	public void checkSavingSalesInvoiceVoucherRmaItemWithHYDWarehouse()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingSalesInvoiceVoucherRmaItemWithHYDWarehouse(), true);
	}
	
	@Test(priority = 35)   // getting issue in  Rma item  Overall avg Rate not matching.
	public void checkSavingSalesInvoiceVoucherRmaItemWithSECWarehouse()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingSalesInvoiceVoucherRmaItemWithSECWarehouse(), true);
	}
	
	@Test(priority = 36)
	public void checkStockLedgerReportForBrBatchItemAfterSavingSalesInvoiceVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForBrBatchItemAfterSavingSalesInvoiceVocher(), true);
	}
	
	@Test(priority = 38)
	public void checkStockLedgerReportForRmaItemAfterSavingSalesInvoiceVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForRmaItemAfterSavingSalesInvoiceVocher(), true);
	}
	
	//Stock Transfer voucher1
	
	@Test(priority = 39)  // getting issue in  Rma item  Overall avg Rate not matching.
	public void checkLoginToSUUserAndSavingStockTransferVocher1()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkLoginToSUUserAndSavingStockTransferVocher1(), true);
	}
	
	@Test(priority = 40)
	public void validteItemQueryStocksInBatchItemAfterSavingStockTransferVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInBatchItemAfterSavingStockTransferVoucher(), true);
	}
	
	@Test(priority = 41)
	public void validteItemQueryStocksInFifoItemAfterSavingStockTransferVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInFifoItemAfterSavingStockTransferVoucher(), true);
	}
	
	@Test(priority = 42)
	public void validteItemQueryStocksInWAItemAfterSavingStockTransferVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInWAItemAfterSavingStockTransferVoucher(), true);
	}
	
	@Test(priority = 43)
	public void checkStockLedgerReportForBrBatchItemAfterStockTransferVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForBrBatchItemAfterStockTransferVoucher(), true);
	}
	
	@Test(priority = 44) 
	public void checkStockLedgerReportForFifoItemAfterStockTransferVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForFifoItemAfterStockTransferVoucher(), true);
	}
	
	@Test(priority = 45)
	public void checkStockLedgerReportForRmaItemAfterStockTransferVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForRmaItemAfterStockTransferVoucher(), true);
	}
	
	// Stock Transfer voucher2
	
	@Test(priority = 46)  // getting issue in Rma item  Overall avg Rate not matching.
	public void checkSavingStockTransferVocher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingStockTransferVocher2(), true);
	}
	
	@Test(priority = 47)
	public void validteItemQueryStocksInBatchItemAfterSavingStockTransferVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInBatchItemAfterSavingStockTransferVoucher2(), true);
	}
	
	@Test(priority = 48)  
	public void validteItemQueryStocksInFifoItemAfterSavingStockTransferVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInFifoItemAfterSavingStockTransferVoucher2(), true);
	}
	
	@Test(priority = 49) 
	public void validteItemQueryStocksInWAItemAfterSavingStockTransferVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInWAItemAfterSavingStockTransferVoucher2(), true);
	}
	
	@Test(priority = 50) 
	public void checkStockLedgerReportForBrBatchItemAfterStockTransferVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForBrBatchItemAfterStockTransferVoucher2(), true);
	}
	
	@Test(priority = 51) 
	public void checkStockLedgerReportForFifoItemAfterStockTransferVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForFifoItemAfterStockTransferVoucher2(), true);
	}
	
	@Test(priority = 52) 
	public void checkStockLedgerReportForRmaItemAfterStockTransferVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForRmaItemAfterStockTransferVoucher2(), true);
	}
	
	@Test(priority = 53) // getting issue in  Rma item  Overall avg Rate not matching.
	public void checkSavingexcessInStocksVoucherWithALLItemsInInvUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingexcessInStocksVoucherWithALLItemsInInvUser(), true);
	}
	
	@Test(priority = 54)
	public void checkSavingShortagesinStockVoucherWithBrItemHydSecWareHosueInSUUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingShortagesinStockVoucherWithBrItemHydSecWareHosueInSUUser(), true);
	}
	
	@Test(priority = 55)
	public void checkSavingShortagesinStockVoucherWithBrItemTestWareHosue()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingShortagesinStockVoucherWithBrItemTestWareHosue(), true);
	}
	
	@Test(priority = 56)
	public void checkSavingShortagesinStockVoucherWithFifoItem3WareHosue()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingShortagesinStockVoucherWithFifoItem3WareHosue(), true);
	}
	
	@Test(priority = 57) // getting issue in  Rma item  Overall avg Rate not matching.
	public void checkSavingShortagesinStockVoucherWithRmaItemHydWareHosue()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingShortagesinStockVoucherWithRmaItemHydWareHosue(), true);
	}
	
	@Test(priority = 58) // getting issue in  Rma item  Overall avg Rate not matching.
	public void checkSavingShortagesinStockVoucherWithRmaItemSecWareHosue()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingShortagesinStockVoucherWithRmaItemSecWareHosue(), true);
	}
	
	@Test(priority = 59) // getting issue in  Rma item  Overall avg Rate not matching.
	public void checkSavingShortagesinStockVoucherWithRmaItemTestWareHosue()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingShortagesinStockVoucherWithRmaItemTestWareHosue(), true);
	}
	
	@Test(priority = 60) 
	public void checkSavingPurchaseReturnsVoucherWithBrItemHydSecWareHosueInINVUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingPurchaseReturnsVoucherWithBrItemHydSecWareHosueInINVUser(), true);
	}
	
	@Test(priority = 61) 
	public void checkSavingPurchaseReturnsVoucherWithFifoItemHydSecWareHosueInINVUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingPurchaseReturnsVoucherWithFifoItemHydSecWareHosueInINVUser(), true);
	}
	
	@Test(priority = 62)  // getting issue in  Rma item  Overall avg Rate not matching.
	public void checkSavingPurchaseReturnsVoucherWithRmaItemHydWareHosueInINVUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingPurchaseReturnsVoucherWithRmaItemHydWareHosueInINVUser(), true);
	}
	
	@Test(priority = 63)  // getting issue in  Rma item  Overall avg Rate not matching.
	public void checkSavingPurchaseReturnsVoucherWithRmaItemSECWareHosueInINVUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingPurchaseReturnsVoucherWithRmaItemSECWareHosueInINVUser(), true);
	}
	
	@Test(priority = 64)  
	public void checkSavingSalesReturnsVoucherWithAllItemsINVUser()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingSalesReturnsVoucherWithAllItemsINVUser(), true);
	}
	
	@Test(priority = 65)  
	public void validteItemQueryStocksInBatchItemAfterSavingSalesReturnsVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInBatchItemAfterSavingSalesReturnsVocher(), true);
	}
	
	@Test(priority = 66)  
	public void validteItemQueryStocksInFifoItemAfterSavingSalesReturnsVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInFifoItemAfterSavingSalesReturnsVocher(), true);
	}
	
	@Test(priority = 67)  
	public void validteItemQueryStocksInWAItemAfterSavingSalesReturnsVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.validteItemQueryStocksInWAItemAfterSavingSalesReturnsVocher(), true);
	}
	
	@Test(priority = 68)  
	public void checkStockLedgerReportForBrBatchItemAfterSavingSalesReturnsVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForBrBatchItemAfterSavingSalesReturnsVocher(), true);
	}
	
	//@Test(priority = 69)  // values taken from item query 
	public void checkStockLedgerReportForFifoItemAfterSavingSalesReturnVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForFifoItemAfterSavingSalesReturnVocher(), true);
	}
	
	@Test(priority = 70)  
	public void checkStockLedgerReportForRmaItemAfterSavingSalesReturnvoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForRmaItemAfterSavingSalesReturnvoucher(), true);
	}
	
	@Test(priority = 71)  //header checked
	public void checkStockMovementReportByDefault()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockMovementReportByDefault(), true);
	}
	
	@Test(priority = 72)  
	public void checkStockMovementReportByWareHouse()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockMovementReportByWareHouse(), true);
	}
	
	@Test(priority = 73)  //header checked
	public void checkStockValuationReportByProduct()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockValuationReportByProduct(), true);
	}
	
	@Test(priority = 74)  
	public void checkStockValuationReportByProductbyInventoryTag()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockValuationReportByProductbyInventoryTag(), true);
	}
	
	@Test(priority = 75)  
	public void checkStockValuationReportByInventoryTagByProduct()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockValuationReportByInventoryTagByProduct(), true);
	}
	
	@Test(priority = 76)  //header checked
	public void checkStockAgeingAnalysisReportByDefault()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockAgeingAnalysisReportByDefault(), true);
	}
	
	@Test(priority = 77)  
	public void checkStockAgeingAnalysisReportByWarehouse()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockAgeingAnalysisReportByWarehouse(), true);
	}
	
	@Test(priority = 78)  
	public void checkStockAgeingAnalysisReportByWarehouseHyd()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockAgeingAnalysisReportByWarehouseHyd(), true);
	}
	
	@Test(priority = 79)  
	public void checkStockAgeingAnalysisReportByWarehouseSEC()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockAgeingAnalysisReportByWarehouseSEC(), true);
	}
	
	@Test(priority = 80)  //header checked
	public void checkStockBalanceByWarehouseReport()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockBalanceByWarehouseReport(), true);
	}
	
	@Test(priority = 81)  
	public void checkStockStatementReporFortBrBatchItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockStatementReporFortBrBatchItem(), true);
	}
	
	@Test(priority = 82)  
	public void checkStockStatementReportForFifoItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockStatementReportForFifoItem(), true);
	}
	
	@Test(priority = 83)  
	public void checkStockStatementReportForRmaItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockStatementReportForRmaItem(), true);
	}
	
	
	@Test(priority = 84)  //Select all check box
	public void checkStockLedgerReportAfterSavingSalesReturnsVocherWithRefineFilterFirstPage()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportAfterSavingSalesReturnsVocherWithRefineFilterFirstPage(), true);
	}
	
	@Test(priority = 85)   //Select all check box
	public void checkStockLedgerReportAfterSavingSalesReturnsVocherWithRefineFilterLastPage()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportAfterSavingSalesReturnsVocherWithRefineFilterLastPage(), true);
	}
	
	
	@Test(priority = 86)  //Select all check box
	public void checkStockMovementReportWithRefineFilter()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockMovementReportWithRefineFilter(), true);
	}
	
	
	@Test(priority = 87)  //Select all check box
	public void checkStockValuationReportByProductWithRefineFilter()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockValuationReportByProductWithRefineFilter(), true);
	}
	
	
	@Test(priority = 88)  //Select all check box
	public void checkStockAgeingAnalysisReportByDefaultWithRefineFilter()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockAgeingAnalysisReportByDefaultWithRefineFilter(), true);
	}
	
	@Test(priority = 89)  //Select all check box
	public void checkStockBalanceByWarehouseReportWithRefineFilter()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockBalanceByWarehouseReportWithRefineFilter(), true);
	}
	
	@Test(priority = 90)  //Select all check box
	public void checkLedgerReportWithRefineFilterFirstPage()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkLedgerReportWithRefineFilterFirstPage(), true);
	}
	
	@Test(priority = 91)  //Select all check box
	public void checkLedgerReportWithRefineFilterLastPage()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkLedgerReportWithRefineFilterLastPage(), true);
	}
	
	@Test(priority = 92)  
	public void checkPurchaseRegisterReportWithRefineFilter()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkPurchaseRegisterReportWithRefineFilter(), true);
	}
	
	@Test(priority = 93)  
	public void checkTrailBalanceReportWithRefineFilter()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkTrailBalanceReportWithRefineFilter(), true);
	}
	
	@Test(priority = 94)  //Select all check box
	public void checkCustomerAgeingSummeryReportWithRefineFilter()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkCustomerAgeingSummeryReportWithRefineFilter(), true);
	}
	
	@Test(priority = 95)  //Select all check box
	public void checkVendorAgeingSummeryReportWithRefineFilter()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkVendorAgeingSummeryReportWithRefineFilter(), true);
	}
	
	
	
	// Stock In Stock Out Vouchers
	@Test(priority = 101)  
	public void checkSavingStockOutVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingStockOutVoucher(), true);
	}
	
	@Test(priority = 102)  
	public void checkSavingStockInVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkSavingStockInVoucher(), true);
	}
	
	@Test(priority = 103)  
	public void checkStockLedgerReportForStockItemAfterStockInAndStockOutVouchers()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForStockItemAfterStockInAndStockOutVouchers(), true);
	}
	
	@Test(priority = 104)  
	public void checkLoginToTest1UserAndAuthorizeStockOutVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkLoginToTest1UserAndAuthorizeStockOutVocher(), true);
	}
	
	@Test(priority = 105)  
	public void checkStockLedgerReportForStockItemAfterAutrhroizedStockOutVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		ip = new ItemQueryPage(getDriver());
		Assert.assertEquals(ip.checkStockLedgerReportForStockItemAfterAutrhroizedStockOutVoucher(), true);
	} 
	
	

}
