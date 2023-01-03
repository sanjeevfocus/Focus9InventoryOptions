package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.TransactionsPage;
import com.focus.base.BaseEngine;

public class TransactionsTest extends BaseEngine {
	TransactionsPage TP;

	// Create View

	@Test(priority = 385)
	public void checkSavingCreateViewInHomePageOfPurchaseVoucherVAT()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingCreateViewInHomePageOfPurchaseVoucherVAT(), true);
	}

	@Test(priority = 386)
	public void checkEditingCreatedViewInHomePageOfPurchaseVoucherVAT()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkEditingCreatedViewInHomePageOfPurchaseVoucherVAT(), true);
	}

	@Test(priority = 387)
	public void checkSavingCreateViewInHomePageForDelete()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingCreateViewInHomePageForDelete(), true);
	}

	@Test(priority = 388)
	public void checkDeletingCreatedViewInHomePageOfPurchaseVoucherVAT()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkDeletingCreatedViewInHomePageOfPurchaseVoucherVAT(), true);
	}

	// Customization

	@Test(priority = 389)
	public void checkVoucherHomeCustomizationAddingDepartment()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkVoucherHomeCustomizationAddingDepartment(), true);
	}

	@Test(priority = 390)
	public void checkVoucherHomeCustomizationEditingAndAddingWarehouse()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkVoucherHomeCustomizationEditingAndAddingWarehouse(), true);
	}

	@Test(priority = 391)
	public void checkVoucherHomeCustomizationEditingAndRemoveWarehouse()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkVoucherHomeCustomizationEditingAndRemoveWarehouse(), true);
	}

	@Test(priority = 400)
	public void checkCustomizingFieldForVendorAccountInPurchaseVoucherVAT()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkCustomizingFieldForVendorAccountInPurchaseVoucherVAT(), true);
	}

	// Purchase Voucher VAT

	@Test(priority = 401)
	public void checkSavingPurchaseVoucherVATVoucher1WithAllItems()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingPurchaseVoucherVATVoucher1WithAllItems(), true);
	}

	@Test(priority = 402)
	public void checkEditingPurchaseVoucherVATVoucher1()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkEditingPurchaseVoucherVATVoucher1(), true);
	}

	@Test(priority = 403)
	private void checkSavingPurchaseVoucherVATVoucher2WithAllItems()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingPurchaseVoucherVATVoucher2WithAllItems(), true);
	}

	@Test(priority = 404)
	private void checkDeletingPurchaseVoucherVATVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkDeletingPurchaseVoucherVATVoucher2(), true);
	}

	// Update MRN stock value through purchase

	@Test(priority = 409)
	private void checkEnablingUpdateMRNStockValueThroughPurchase()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkEnablingUpdateMRNStockValueThroughPurchase(), true);
	}

	@Test(priority = 410)
	private void checkSavingDesignWorkflowFromMRNToPurchasesVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingDesignWorkflowFromMRNToPurchasesVoucher(), true);
	}

	@Test(priority = 411)
	private void checkSavingMaterialReceiptNotesVoucher1WithStockItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingMaterialReceiptNotesVoucher1WithStockItem(), true);
	}

	@Test(priority = 412)
	private void checkSavingPurchaseVoucherVoucher1WithoutUpdateStock()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingPurchaseVoucherVoucher1WithoutUpdateStock(), true);
	}

	@Test(priority = 413)
	private void checkValidatingMaterialReceiptNotesVoucher1AfterSavingPurchasesVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkValidatingMaterialReceiptNotesVoucher1AfterSavingPurchasesVoucher(), true);
	}

	@Test(priority = 414)
	private void checkSavingMaterialReceiptNotesVoucher2WithStockItem()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingMaterialReceiptNotesVoucher2WithStockItem(), true);
	}

	@Test(priority = 415)
	private void checkCreatingSChargesFieldInPurchaseVoucherBody()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkCreatingSChargesFieldInPurchaseVoucherBody(), true);
	}

	@Test(priority = 416)
	private void checkSavingPurchaseVoucherVoucher2WithSChargesFieldByConvertingPendingLink()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingPurchaseVoucherVoucher2WithSChargesFieldByConvertingPendingLink(), true);
	}

	@Test(priority = 417)
	private void checkValidatingMaterialReceiptNotesVoucher2AfterSavingPurchasesVoucher2WithSCharges()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkValidatingMaterialReceiptNotesVoucher2AfterSavingPurchasesVoucher2WithSCharges(),
				true);
	}

	@Test(priority = 418)
	private void checkCreatingMChargesFieldInMaterialReceiptNotesVoucherBody()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkCreatingMChargesFieldInMaterialReceiptNotesVoucherBody(), true);
	}

	@Test(priority = 419)
	private void checkSavingMaterialReceiptNotesVoucher3WithMChargesField()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingMaterialReceiptNotesVoucher3WithMChargesField(), true);
	}

	@Test(priority = 420)
	private void checkLoadingMChargesFieldInPurchaseVoucherBodyFromMRN()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkLoadingMChargesFieldInPurchaseVoucherBodyFromMRN(), true);
	}

	@Test(priority = 421)
	private void checkSavingPurchaseVoucher3WithSChargesAndMChargesFields()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingPurchaseVoucher3WithSChargesAndMChargesFields(), true);
	}

	@Test(priority = 422)
	private void checkValidatingMaterialReceiptNotesVoucher3AfterSavingPurchasesVoucher3WithSChargesAndMCharges()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(
				TP.checkValidatingMaterialReceiptNotesVoucher3AfterSavingPurchasesVoucher3WithSChargesAndMCharges(),
				true);
	}

	// Sales Invoice VAT

	@Test(priority = 430)
	private void checkSavingSalesInvoiceVATVoucher1WithAllItems()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingSalesInvoiceVATVoucher1WithAllItems(), true);
	}

	@Test(priority = 431)
	private void checkEditingSalesInvoiceVATVoucher1AndUpdating()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkEditingSalesInvoiceVATVoucher1AndUpdating(), true);
	}

	@Test(priority = 432)
	private void checkSavingSalesInvoiceVATVoucher2WithAllItems()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingSalesInvoiceVATVoucher2WithAllItems(), true);
	}

	@Test(priority = 433)
	private void checkDeletingSalesInvoiceVoucherVATVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkDeletingSalesInvoiceVoucherVATVoucher2(), true);
	}

	@Test(priority = 434)
	private void checkPrintPdfOfSalesInvoiceVATVoucher1() throws AWTException, Exception {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkPrintPdfOfSalesInvoiceVATVoucher1(), true);
	}

	
	@Test(priority = 435)
	private void checkEmailOptionInSalesInvoiceVATVoucher1() throws AWTException, Exception {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkEmailOptionInSalesInvoiceVATVoucher1(), true);
	}

	// Receipts

	@Test(priority = 440)
	private void checkSavingRecepitsVATVoucher1()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingRecepitsVATVoucher1(), true);
	}

	@Test(priority = 441)
	private void checkEditingRecepitsVATVoucher1()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkEditingRecepitsVATVoucher1(), true);
	}

	@Test(priority = 442)
	private void checkSavingRecepitsVATVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingRecepitsVATVoucher2(), true);
	}

	@Test(priority = 443)
	private void checkDeletingRecepitsVATVocher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkDeletingRecepitsVATVocher2(), true);
	}

	// Payments

	@Test(priority = 448)
	private void checkSavingPaymentVATVoucher1()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingPaymentVATVoucher1(), true);
	}

	@Test(priority = 449)
	private void checkEditingPaymentVATVoucher1()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkEditingPaymentVATVoucher1(), true);
	}

	@Test(priority = 450)
	private void checkSavingPaymentVATVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingPaymentVATVoucher2(), true);
	}

	@Test(priority = 451)
	private void checkDeletingPaymentVATVoucher2()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkDeletingPaymentVATVoucher2(), true);
	}

	// PDC

	@Test(priority = 460)
	private void checkSavingVoucherINPDR()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingVoucherINPDR(), true);
	}

	@Test(priority = 461)
	private void chekEditingVoucherINPDRVAT()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.chekEditingVoucherINPDRVAT(), true);
	}

	@Test(priority = 462)
	private void checkDeletingVoucherInPDRVAT()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkDeletingVoucherInPDRVAT(), true);
	}

	@Test(priority = 463)
	private void checkSavingVoucherINPDPVAT()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkSavingVoucherINPDPVAT(), true);
	}

	@Test(priority = 464)
	private void chekEditingVoucherINPDPVAT()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.chekEditingVoucherINPDPVAT(), true);
	}

	@Test(priority = 465)
	private void checkAddExtraFieldPDCInHeaderFieldOFPayments()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkAddExtraFieldPDCInHeaderFieldOFPayments(), true);
	}

	@Test(priority = 466)
	private void checkAddExtraFieldPDCInHeaderFieldOFReceipts()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkAddExtraFieldPDCInHeaderFieldOFReceipts(), true);
	}

	@Test(priority = 467)
	private void checkConveringVouchersINPDCMaturedScreen()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkConveringVouchersINPDCMaturedScreen(), true);
	}

	@Test(priority = 468)
	private void checkConvertedVoucherInPayments()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkConvertedVoucherInPayments(), true);
	}

	@Test(priority = 469)
	private void checkConvertedVoucherInRecepits()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		TP = new TransactionsPage(getDriver());
		Assert.assertEquals(TP.checkConvertedVoucherInRecepits(), true);
	}

}
