package com.focus.Pages;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.Color;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.focus.base.BaseEngine;

public class ItemQueryPage extends BaseEngine {

	public ItemQueryPage(WebDriver driver) {
		PageFactory.initElements(driver, this);

	}

	@FindBy(xpath = "//div[@class='theme_color font-6']")
	public static WebElement errorMessage;

	@FindBy(xpath = "//span[@class='icon-reject2 theme_color']")
	public static WebElement errorMessageCloseBtn;

	public static String checkValidationMessage(String ExpMessage)
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		try {
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
			String actErrorMessage = errorMessage.getText();
			String expErrorMessage = ExpMessage;

			try {
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
				errorMessageCloseBtn.click();

				System.out.println("ValidationMessage  :  " + actErrorMessage + " Value Expected : " + expErrorMessage);

				return actErrorMessage;
			} catch (Exception ee) {

				System.out.println("ValidationMessage  :  " + actErrorMessage + " Value Expected : " + expErrorMessage);

				return actErrorMessage;
			}
		} catch (Exception e) {
			System.err.println("Error Message NOT Found or NOT Clickable");
			System.err.println(e.getMessage());

			String Exception = e.getMessage();

			return Exception;
		}
	}

	@FindBy(xpath = "//a[@id='135']/div")
	private static WebElement inventoryMenu;
	
	@FindBy(xpath = "//a[@id='137']/span")
	private static WebElement inventoryTransactionsMenu;
	
	@FindBy(xpath = "//a[@id='2033']/span")
	private static WebElement inventoryTransactionsStockMenu;
	
	@FindBy(xpath = "//a[@id='2056']/span")
	private static WebElement inventoryTransactionsStockOut;
	
	@FindBy(xpath="//*[@id='id_transaction_viewcontainer']/li[1]/a")
	private static WebElement  homePageAllVouchersView;
	
	@FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[13]")
	private static WebElement voucherGridRow1AuthorizationStatus;
	
	@FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[2]/td[13]")
	private static WebElement voucherGridRow2AuthorizationStatus;
	
	@FindBy(xpath = "//a[@id='2057']/span")
	private static WebElement inventoryTransactionsStockIn;
	
	@FindBy(xpath = "//a[@id='2036']/span")
	private static WebElement inventoryTransactionsStockMenuStockTransfersMenu;

	@FindBy(xpath = "//a[@id='200']/span")
	private static WebElement InventoryReportsMenu;

	@FindBy(xpath = "//a[@id='202']/span")
	private static WebElement itemqueryMenu;

	@FindBy(xpath = "//input[@id='productQuery']")
	private static WebElement itemTxt;
	
	@FindBy(xpath = "//div[@id='id_transaction_homescreen_authorize']")
	private static WebElement stockOutHomeAuthBtn;

	@FindBy(xpath = "//a[@class='transaction_viewname_anchor font-5']")
	private static WebElement stockOutPendingAuthLink;
	
	@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr/td[5]")
	private static List<WebElement> grid_VoucherNoList;
	
	@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr/td[2]/input")
	private static List<WebElement> grid_CheckBoxList;
	
	@FindBy(xpath = "(//button[@class='btn btn-default btn-sm'])[1]")
	private static WebElement getStockBtn;

	@FindBy(xpath = "//span[@id='currentStock']")
	private static WebElement currentStockTxt;

	@FindBy(xpath = "//span[@id='avgStockRate']")
	private static WebElement avgStockRateTxt;

	@FindBy(xpath = "//span[@id='value']")
	private static WebElement stockValueTxt;

	@FindBy(xpath = "//span[@id='openingStock']")
	private static WebElement openStockQuantityTxt;

	@FindBy(xpath = "(//span[@id='ProdQuery_ShowPopup'])[1]")
	private static WebElement balanceByWarehouseBtn;

	@FindBy(xpath = "//*[@id='tagWiseData']/tbody/tr[1]/td")
	private static List<WebElement> row1BalanceByWareHouseList;

	@FindBy(xpath = "//*[@id='tagWiseData']/tbody/tr[2]/td")
	private static List<WebElement> row2BalanceByWareHouseList;

	@FindBy(xpath = "//*[@id='tagWiseData']/tbody/tr[3]/td")
	private static List<WebElement> row3BalanceByWareHouseList;
	
	@FindBy(xpath = "//*[@id='tagWiseData']/tbody/tr[4]/td")
	private static List<WebElement> row4BalanceByWareHouseList;

	public boolean validteItemQueryStocksInBatchItem() throws InterruptedException {

		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(itemqueryMenu);
		click(itemTxt);
		Thread.sleep(2000);
		enterText(itemTxt, "BR-BATCH");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);
		
		String actlOpenStockQuantity  = getText(openStockQuantityTxt);
		String expOpenStockQuantity   = "20.000";

		String actlCurrentStock       = getText(currentStockTxt);
		String expCurrentStock        = "5.000";

		String actlavgStockRate       = getText(avgStockRateTxt);
		String expavgStockRate        = "8.0626";

		String actlstockValueTxt      = getText(stockValueTxt);
		String expstockValueTxt       = "40.3131";
		

		System.out.println("=====Stocks In Batch Items======");

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock           :" + actlCurrentStock);
		System.out.println("expCurrentStock            :" + expCurrentStock);

		System.out.println("actlavgStockRate           :" + actlavgStockRate);
		System.out.println("expavgStockRate            :" + expavgStockRate);

		System.out.println("actlstockValueTxt          :" + actlstockValueTxt);
		System.out.println("expstockValueTxt           :" + expstockValueTxt);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)) {
			return true;
		} else {
			return false;
		}

	}

	public boolean validteItemQueryStocksInFifoItem() throws InterruptedException {

		click(itemTxt);
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.BACK_SPACE);
		enterText(itemTxt, "FIFO-BIN");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "5.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "8.2161";

		String actlstockValueTxt     = getText(stockValueTxt);
		String expstockValueTxt      = "41.0807";

		System.out.println("=====Stocks In FIFO Items======");

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock           :" + actlCurrentStock);
		System.out.println("expCurrentStock            :" + expCurrentStock);

		System.out.println("actlavgStockRate           :" + actlavgStockRate);
		System.out.println("expavgStockRate            :" + expavgStockRate);

		System.out.println("actlstockValueTxt          :" + actlstockValueTxt);
		System.out.println("expstockValueTxt           :" + expstockValueTxt);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)) {
			return true;
		} else {
			return false;
		}

	}

	public boolean validteItemQueryStocksInWAItem() throws InterruptedException {

		click(itemTxt);
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.BACK_SPACE);
		enterText(itemTxt, "WA-RMA");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "5.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "9.6019";

		String actlstockValueTxt     = getText(stockValueTxt);
		String expstockValueTxt      = "48.0097";

		ScrollToElement(balanceByWarehouseBtn);
		click(balanceByWarehouseBtn);

		Thread.sleep(3000);

		int count = row1BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow1 = new ArrayList<String>();

		for (int i = 0; i < count; i++) {
			String data = row1BalanceByWareHouseList.get(i).getText();
			warehouseRow1.add(data);
		}

		String actwarehouseRow1List = warehouseRow1.toString();
		String expWarehouseRow1List = "[1, HYD, 2.000, 0.000, 2.000, 9.4927, 18.9854]";

		int count1 = row2BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow2 = new ArrayList<String>();

		for (int i = 0; i < count1; i++) {
			String data = row2BalanceByWareHouseList.get(i).getText();
			warehouseRow2.add(data);
		}

		String actwarehouseRow2List = warehouseRow2.toString();
		String expWarehouseRow2List = "[2, SEC, 1.000, 0.000, 1.000, 9.0458, 9.0458]";

		int count2 = row3BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow3 = new ArrayList<String>();

		for (int i = 0; i < count2; i++) {
			String data = row3BalanceByWareHouseList.get(i).getText();
			warehouseRow3.add(data);
		}

		String actwarehouseRow3List = warehouseRow3.toString();
		String expWarehouseRow3List = "[3, TEST, 2.000, 0.000, 2.000, 11.3062, 22.6125]";

		System.out.println("=====Stocks In WA Items======");
		
		

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock           :" + actlCurrentStock);
		System.out.println("expCurrentStock            :" + expCurrentStock);

		System.out.println("actlavgStockRate           :" + actlavgStockRate);
		System.out.println("expavgStockRate            :" + expavgStockRate);

		System.out.println("actlstockValueTxt          :" + actlstockValueTxt);
		System.out.println("expstockValueTxt           :" + expstockValueTxt);

	
		
		System.out.println("===========Balance By Warehouse=======");

		System.out.println(actwarehouseRow1List);
		System.out.println(expWarehouseRow1List);

		System.out.println(actwarehouseRow2List);
		System.out.println(expWarehouseRow2List);

		System.out.println(actwarehouseRow3List);
		System.out.println(expWarehouseRow3List);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)
				&& actwarehouseRow1List.equalsIgnoreCase(expWarehouseRow1List)
				&& actwarehouseRow2List.equalsIgnoreCase(expWarehouseRow2List)
				&& actwarehouseRow3List.equalsIgnoreCase(expWarehouseRow3List)) {
			return true;
		} else {
			return false;
		}

	}

	@FindBy(xpath = "//*[@id='selectAllMasters_']")
	private static WebElement selectAllItemsCheckBox;
	
	@FindBy(xpath = "//a[@id='551']/span")
	private static WebElement InventoryStockLedgerMenu;

	@FindBy(xpath = "//tbody[@id='LandingGridBody']/tr/td[8]/div/label/input")
	private static List<WebElement> stockLedgerHometableItemChkboxList;

	@FindBy(xpath = "//*[@id='LandingGridBody']/tr/td[12]")
	private static List<WebElement> ledgerHometableItemNamesList;

	@FindBy(xpath = "//*[@id='LandingGridBody']/tr")
	private static List<WebElement> stockLedgerHometableRowCount;

	@FindBy(xpath = "//select[@id='DateOptions_']")
	private static WebElement sl_DateOptionDropdown;
	
	@FindBy(xpath = "//*[@id='MasterGroup__101']")
	private static WebElement accountDropdown;

	@FindBy(xpath = "//i[contains(@class,'icon-font6 icon-ok')]")
	private static WebElement sl_OkBtn;

	@FindBy(xpath = "//button[contains(text(), 'Last')]")
	private static WebElement stockLedgerLastBtn;
	
	@FindBy(xpath = "//button[contains(text(), 'Previous')]")
	private static WebElement stockLedgerPreviousBtn;

	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[9]")
	private static WebElement stockLedgerBrBatchBalanceQty;
	
	@FindBy(xpath = "(//tr[@id='trRender_5'])[1]/td[9]")
	private static WebElement stockLedgerBrBatchAfterSavingpurchaseBalanceQty;
	
	@FindBy(xpath = "(//tr[@id='trRender_2'])[1]/td[9]")
	private static WebElement stockLedgerBrBatchAfterSavingSalesInvoiceBalanceQty;
	
	@FindBy(xpath = "(//tr[@id='trRender_2'])[1]/td[12]")
	private static WebElement stockLedgerBrBatchAfterSavingSalesInvoiceBalanceValue;
	
	@FindBy(xpath = "(//tr[@id='trRender_2'])[1]/td[13]")
	private static WebElement stockLedgerBrBatchAfterSavingSalesInvoiceAvgRate;
	
	@FindBy(xpath = "(//tr[@id='trRender_10'])[1]/td[8]")
	private static WebElement stockLedgerBrBatchAfterSavingSalesInvoiceIssuedrateTxt;
	
	@FindBy(xpath = "(//tr[@id='trRender_4'])[1]/td[8]")
	private static WebElement stockLedgerFifoItemAfterSavingSalesInvoiceIssuedrateTxt;
	
	@FindBy(xpath = "(//tr[@id='trRender_8'])[1]/td[8]")
	private static WebElement stockLedgerRmaItemAfterSavingSalesInvoiceIssuedrateTxt;
	
	@FindBy(xpath = "(//tr[@id='trRender_5'])[1]/td[9]")
	private static WebElement stockLedgerBrBatchBalanceQtyAfterLoginToInvUser;
	
	@FindBy(xpath = "(//tr[@id='trRender_5'])[1]/td[9]")
	private static WebElement stockLedgerFifoItemBalanceQtyAfterLoginToInvUser;
	
	@FindBy(xpath = "(//tr[@id='trRender_5'])[1]/td[12]")
	private static WebElement stockLedgerFifoItemBalancevaluAfterLoginToInvUser;
	
	@FindBy(xpath = "(//tr[@id='trRender_5'])[1]/td[13]")
	private static WebElement stockLedgerFifoItemAvgRateAfterLoginToInvUser;


	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[12]")
	private static WebElement stockLedgerBrBatchBalanceValue;
	
	@FindBy(xpath = "(//tr[@id='trRender_5'])[1]/td[12]")
	private static WebElement stockLedgerBrBatchBalanceValueAfterSavingpurchase;
	
	@FindBy(xpath = "(//tr[@id='trRender_5'])[1]/td[12]")
	private static WebElement stockLedgerBrBatchBalanceValueAfterLoginToInvUser;

	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[13]")
	private static WebElement stockLedgerBrBatchAvgRate;
	
	@FindBy(xpath = "(//tr[@id='trRender_5'])[1]/td[13]")
	private static WebElement stockLedgerBrBatchAvgRateAfterSavingpurchase;	
	
	@FindBy(xpath = "(//tr[@id='trRender_5'])[1]/td[13]")
	private static WebElement stockLedgerBrBatchAvgRateAfterLoginToInvUser;
	
	
	@FindBy(xpath = "(//tr[@id='trRender_7'])[1]/td[9]")
	private static WebElement stockLedgerFifoItemBalanceQty;
	
	@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[9]")
	private static WebElement stockLedgerFifoItemBalanceQtyAfterSavingPurchase;


	@FindBy(xpath = "(//tr[@id='trRender_7'])[1]/td[12]")
	private static WebElement stockLedgerFifoItemBalanceValue;
	
	@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[12]")
	private static WebElement stockLedgerFifoItemBalanceValueAfterSavingPurchase;

	@FindBy(xpath = "(//tr[@id='trRender_7'])[1]/td[13]")
	private static WebElement stockLedgerFifoItemAvgRate;
	
	@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[13]")
	private static WebElement stockLedgerFifoItemAvgRateAfterSavingPurchase;

	@FindBy(xpath = "(//tr[@id='trRender_8'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemBalanceQty;
	
	@FindBy(xpath = "(//tr[@id='trRender_10'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemBalanceQtyAftersavingPurchase;
	
	@FindBy(xpath = "(//tr[@id='trRender_6'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemBalanceQtyAfterLoginToInvUser;

	@FindBy(xpath = "(//tr[@id='trRender_8'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemBalanceValue;
	
	@FindBy(xpath = "(//tr[@id='trRender_10'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemBalanceValueAfterSavingPurchase;
	
	@FindBy(xpath = "(//tr[@id='trRender_6'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemBalanceValueAfterLoginToInvUser;

	@FindBy(xpath = "(//tr[@id='trRender_8'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemAvgRate;
	
	@FindBy(xpath = "(//tr[@id='trRender_10'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemAvgRateAfterSavingPurchase;
	
	@FindBy(xpath = "(//tr[@id='trRender_6'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemAvgRateAfterLogintoInvUser;

	@FindBy(xpath = "(//i[@class='icon-font6 icon-close'])[2]")
	private static WebElement report_CloseBtn;

	@FindBy(xpath = "(//i[contains(@class,'icon-font6 icon-close')])[1]")
	private static WebElement sl_CloseBtn;
	
	
	@FindBy(xpath = "//table[@class='CommonReportTable']//tr//td[3]")
	private static List<WebElement> stockLedgerColCount;
	
	

	public boolean checkStockLedgerReportForBrBatchItem()	
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryStockLedgerMenu);
		Thread.sleep(3000);
		
		int rowcount = stockLedgerHometableRowCount.size();

		System.out.println(rowcount);

		for (int i = 0; i < rowcount; i++) {
			String actName = ledgerHometableItemNamesList.get(i).getText();

			System.out.println(actName);

			if (actName.equalsIgnoreCase("BR-BATCH")) {
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}

		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		
		/*
		
		int rowcount1 = stockLedgerColCount.size();
		 
        System.out.println(rowcount1);

        for (int i = 0; i < rowcount1; i++) {
                String actName = stockLedgerColCount.get(i).getText();

                System.out.println(actName);

                if (actName.equalsIgnoreCase("PurVou:1")) 
                
                {
                	
                	Thread.sleep(6000);
                	
                	//String s1  =stockLedgerColCount.get(i).getCssValue("background-color");
                	String s1=getDriver().findElement(By.xpath("//table[@class='CommonReportTable']//tr//td[3]/..")).getCssValue("background-color");
                	
                	String beforeclick  =stockLedgerColCount.get(i).getCssValue("color");
                	
                	Thread.sleep(2000);
                	
                	String beforeclickHexa = Color.fromString(beforeclick).asHex();
                	  
                	  Thread.sleep(2000);
                	
                	System.err.println("s1   :"+s1);
                	System.err.println("beforeclick   :"+beforeclick);
                	
                	System.err.println("beforeclickHexa   :"+beforeclickHexa);
                	
                	Thread.sleep(2000);
                	
                	
                        stockLedgerColCount.get(i).click();
                      
                        Thread.sleep(6000);
                        
                       // String s3  =stockLedgerColCount.get(i).getCssValue("background-color");
                        
                        String s3=getDriver().findElement(By.xpath("//table[@class='CommonReportTable']//tr//td[3]/..")).getCssValue("background-color");
                    	
                    	String afterclick  =stockLedgerColCount.get(i).getCssValue("color");
                    	Thread.sleep(2000);
                    	
                    	String afterclickHexa = Color.fromString(afterclick).asHex();
                    	
                    	 Thread.sleep(2000);
                    	
                    	System.err.println("s3   :"+s3);
                    	System.err.println("afterclick   :"+afterclick);
                    	System.err.println("afterclickHexa   :"+afterclickHexa);
                    	
                    	if(!afterclickHexa.equalsIgnoreCase("#0073aa"))
                    	{
                    		stockLedgerColCount.get(i).click();
                    		System.err.println("enter into if");
                    	}
                    	
                    	
                    	 
                        
                       
                       
                       Thread.sleep(3000);
                       break;
                        
                }
                
                
        }
        //click(report_CloseBtn);
		return true;
		}


*/
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		click(stockLedgerLastBtn);
		Thread.sleep(3000);
		
		
		
		
		
		
		
		

		String actStockLedgerBrBatchBalanceQty    = getText(stockLedgerBrBatchBalanceQty);
		String expStockLedgerBrBatchBalanceQty    = "5.0000";

		String actStockLedgerBrBatchBalanceValue  = getText(stockLedgerBrBatchBalanceValue);
		String expStockLedgerBrBatchBalanceValue  = "40.3131";

		String actStockLedgerBrBatchAvgRate = getText(stockLedgerBrBatchAvgRate);
		String expStockLedgerBrBatchAvgRate = "8.0626";

		System.out.println("actStockLedgerBrBatchBalanceQty      :" + actStockLedgerBrBatchBalanceQty);
		System.out.println("expStockLedgerBrBatchBalanceQty      :" + expStockLedgerBrBatchBalanceQty);

		System.out.println("actStockLedgerBrBatchBalanceValue    :" + actStockLedgerBrBatchBalanceValue);
		System.out.println("expStockLedgerBrBatchBalanceValue    :" + expStockLedgerBrBatchBalanceValue);

		System.out.println("actStockLedgerBrBatchAvgRate         :" + actStockLedgerBrBatchAvgRate);
		System.out.println("expStockLedgerBrBatchAvgRate         :" + expStockLedgerBrBatchAvgRate);

		if (actStockLedgerBrBatchBalanceQty.equalsIgnoreCase(expStockLedgerBrBatchBalanceQty)
				&& actStockLedgerBrBatchBalanceValue.equalsIgnoreCase(expStockLedgerBrBatchBalanceValue)
				&& actStockLedgerBrBatchAvgRate.equalsIgnoreCase(expStockLedgerBrBatchAvgRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}

	@FindBy(xpath = "//td[contains(text(), 'FIFO-BIN')]//preceding-sibling::td//label/input")
	private static WebElement fifoItemCheckBox;

	@FindBy(xpath = "//td[contains(text(), 'WA-RMA')]//preceding-sibling::td//label/input")
	private static WebElement rmaItemCheckBox;

	@FindBy(xpath = "//td[contains(text(), 'BR-BATCH')]//preceding-sibling::td//label/input")
	private static WebElement brBatchItemCheckBox;

	public boolean checkStockLedgerReportForFifoItem()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == false) {
			click(fifoItemCheckBox);
		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == true) {
			click(rmaItemCheckBox);

		}

		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		
		/*int rowcount2 = stockLedgerColCount.size();

        System.out.println(rowcount2);

        for (int i = 0; i < rowcount2; i++) {
                String actName = stockLedgerColCount.get(i).getText();

                System.out.println(actName);

                if (actName.equalsIgnoreCase("PurVou:1")) 
                
                {
                	
                	Thread.sleep(6000);
                	
                	//String s1  =stockLedgerColCount.get(i).getCssValue("background-color");
                	
                	String s2  =stockLedgerColCount.get(i).getCssValue("color");
                	
                	 String c = Color.fromString(s2).asHex();
                	
                	//System.err.println("s1   :"+s1);
                	System.err.println("s2   :"+s2);
                	
                	System.err.println("c   :"+c);
                	
                	
                        stockLedgerColCount.get(i).click();
                        
                        Thread.sleep(6000);
                        
                       // String s3  =stockLedgerColCount.get(i).getCssValue("background-color");
                    	
                    	String s4  =stockLedgerColCount.get(i).getCssValue("color");
                    	
                    	String c2 = Color.fromString(s4).asHex();
                    	
                    	//System.err.println("s3   :"+s3);
                    	System.err.println("s4   :"+s4);
                    	System.err.println("c2   :"+c2);
                    	
                    	 
                        
                       
                       
                       Thread.sleep(3000);
                       break;
                        
                }
                
                
        }
		return true;
		}*/
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		click(stockLedgerLastBtn);
		Thread.sleep(3000);

		String actstockLedgerFifoItemBalanceQty   = getText(stockLedgerFifoItemBalanceQty);
		String expstockLedgerFifoItemBalanceQty   = "5.0000";

		String actstockLedgerFifoItemBalanceValue = getText(stockLedgerFifoItemBalanceValue);
		String expstockLedgerFifoItemBalanceValue = "41.0807";

		String actstockLedgerFifoItemAvgRate      = getText(stockLedgerFifoItemAvgRate);
		String expstockLedgerFifoItemAvgRate      = "8.2161";

		System.out.println("actstockLedgerFifoItemBalanceQty      :" + actstockLedgerFifoItemBalanceQty);
		System.out.println("expstockLedgerFifoItemBalanceQty      :" + expstockLedgerFifoItemBalanceQty);

		System.out.println("actstockLedgerFifoItemBalanceValue    :" + actstockLedgerFifoItemBalanceValue);
		System.out.println("expstockLedgerFifoItemBalanceValue    :" + expstockLedgerFifoItemBalanceValue);

		System.out.println("actstockLedgerFifoItemAvgRate         :" + actstockLedgerFifoItemAvgRate);
		System.out.println("expstockLedgerFifoItemAvgRate         :" + expstockLedgerFifoItemAvgRate);

		if (actstockLedgerFifoItemBalanceQty.equalsIgnoreCase(expstockLedgerFifoItemBalanceQty)
				&& actstockLedgerFifoItemBalanceValue.equalsIgnoreCase(expstockLedgerFifoItemBalanceValue)
				&& actstockLedgerFifoItemAvgRate.equalsIgnoreCase(expstockLedgerFifoItemAvgRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}  
	
	@FindBy(xpath = "(//i[@class='icon-font6 icon-filter'])[2]")
	private static WebElement stockLedgerFilterBtn;
	
	@FindBy(xpath = "(//div[@class='btn btn-toolbar dropdown theme_button_color'])[1]")
	private static WebElement stockLedgerFilterExpandBtn;
	
	@FindBy(xpath = "//li[text()='Refine']")
	private static WebElement stockLedgerFilterRefineBtn;
	
	@FindBy(xpath = "//*[@id='cmbColumn']")
	private static WebElement filterRefineLookInDropdownBtn;
	
	@FindBy(xpath = "//*[@id='SreachTxt']")
	private static WebElement filterRefineFindWhatTxt;
	
	@FindBy(xpath = "//*[@id='Display']")
	private static WebElement filterRefineLookIn2DropdownBtn;
	
	@FindBy(xpath = "//*[@id='RefineReports']/div/div[3]/div/button[1]")
	private static WebElement filterRefineOkBtn;
	
	@FindBy(xpath = "(//button[text()='Cancel'])[1]")
	private static WebElement filterRefinCancelBtn;
	
	
	
	
	
	
	
	
	@FindBy(xpath = "(//i[@class='icon-font6 icon-filter'])[1]")
	private static WebElement stockLedgerHomePageFilterBtn;
	
	@FindBy(xpath = "//h4[text()='Filter']")
	private static WebElement stockLedgerFilterText;
	
	@FindBy(xpath = "//input[@id='FOption_551_0_DefaultFilter_1']")
	private static WebElement stockLedgerFilterWarehouseField;	
	
	@FindBy(xpath = "//*[@id='FOption_575_0_DefaultFilter_0']")
	private static WebElement stockMovementFilterWarehouseField;
	
	@FindBy(xpath = "//*[@id='FOption_578_0_DefaultFilter_0']")
	private static WebElement stockValuationFilterWarehouseField;
	
	@FindBy(xpath = "//*[@id='FOption_582_0_DefaultFilter_0']")
	private static WebElement stockAgeingFilterWarehouseField;
	
	@FindBy(xpath = "//*[@id='FOption_557_0_DefaultFilter_0']")
	private static WebElement stockStatementFilterWarehouseField;
	
	@FindBy(xpath = "//span[@id='filter_Okbtn_']")
	private static WebElement stockLedgerFilterOkBtn;
	
	@FindBy(xpath = "//span[@id='filterRefresh']")
	private static WebElement stockLedgerFilterRefreshBtn;
	
	@FindBy(xpath = "//*[@id='551_0_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	private static WebElement stockLedgerFilterWarehouseField2;	
	
	@FindBy(xpath = "//*[@id='575_0_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	private static WebElement stockMovementWarehouseField2;
	
	@FindBy(xpath = "//*[@id='578_0_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	private static WebElement stockValuationWarehouseField2;
	
	@FindBy(xpath = "//*[@id='582_0_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	private static WebElement stockAgeingWarehouseField2;
	
	@FindBy(xpath = "//*[@id='557_0_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	private static WebElement stockStatementWarehouseField2;
	
	@FindBy(xpath = "//*[@id='advancefilter_master_575_0_']")
	private static WebElement stockMovementFilterWarehouseField3;
	
	@FindBy(xpath = "//*[@id='advancefilter_master_578_0_']")
	private static WebElement stockValuationFilterWarehouseField3;
	
	@FindBy(xpath = "//*[@id='advancefilter_master_582_0_']")
	private static WebElement stockAgeingFilterWarehouseField3;
	
	@FindBy(xpath = "//*[@id='advancefilter_master_557_0_']")
	private static WebElement stockStatementFilterWarehouseField3;
	
	@FindBy(xpath = "//input[@id='advancefilter_master_551_0_']")
	private static WebElement stockLedgerFilterWarehouseField3;
	
	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemHydBalanceQty;
	
	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemHydBalanceValue;
	
	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemHydAvgRate;
	
	@FindBy(xpath = "(//tr[@id='trRender_2'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemSecBalanceQty;
	
	@FindBy(xpath = "(//tr[@id='trRender_2'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemSecBalanceValue;
	
	@FindBy(xpath = "(//tr[@id='trRender_2'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemSecAvgRate;
	

	public boolean checkStockLedgerReportForRmaItem()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == true) {
			click(fifoItemCheckBox);

		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == false) {
			click(rmaItemCheckBox);

		}

		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(3000);

		String actstockLedgerRmaItemBalanceQty   = getText(stockLedgerRmaItemBalanceQty);
		String expstockLedgerRmaItemBalanceQty   = "5.0000";

		String actstockLedgerRmaItemBalanceValue = getText(stockLedgerRmaItemBalanceValue);
		String expstockLedgerRmaItemBalanceValue = "48.0097";

		String actstockLedgerrmaItemAvgRate      = getText(stockLedgerrmaItemAvgRate);
		String expstockLedgerrmaItemAvgRate      = "9.6019";

		System.out.println("actstockLedgerRmaItemBalanceQty      :" + actstockLedgerRmaItemBalanceQty);
		System.out.println("expstockLedgerRmaItemBalanceQty      :" + expstockLedgerRmaItemBalanceQty);

		System.out.println("actstockLedgerRmaItemBalanceValue    :" + actstockLedgerRmaItemBalanceValue);
		System.out.println("expstockLedgerRmaItemBalanceValue    :" + expstockLedgerRmaItemBalanceValue);

		System.out.println("actstockLedgerrmaItemAvgRate         :" + actstockLedgerrmaItemAvgRate);
		System.out.println("expstockLedgerrmaItemAvgRate         :" + expstockLedgerrmaItemAvgRate);
		
		
		click(stockLedgerFilterBtn);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

		boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
		boolean expFilterpopupIsDisplaying = true;

		System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
				+ expFilterpopupIsDisplaying);
		
		click(stockLedgerFilterWarehouseField);
		enterText(stockLedgerFilterWarehouseField, "HYD");
		Thread.sleep(2000);
		click(stockLedgerFilterOkBtn);
		Thread.sleep(1000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);
		
		String actstockLedgerRmaItemHydBalanceQty   = getText(stockLedgerRmaItemHydBalanceQty);
		String expstockLedgerRmaItemHydBalanceQty   = "2.0000";

		String actstockLedgerRmaItemHydBalanceValue = getText(stockLedgerRmaItemHydBalanceValue);
		String expstockLedgerRmaItemHydBalanceValue = "18.9854";

		String actstockLedgerrmaItemHydAvgRate      = getText(stockLedgerrmaItemHydAvgRate);
		String expstockLedgerrmaItemHydAvgRate      = "9.4927";
		
		System.out.println("=========actstockLedgerRmaItem  HydWareHouse  Values============   : ");
		
		System.out.println("actstockLedgerRmaItemHydBalanceQty      :" + actstockLedgerRmaItemHydBalanceQty);
		System.out.println("expstockLedgerRmaItemHydBalanceQty      :" + expstockLedgerRmaItemHydBalanceQty);

		System.out.println("actstockLedgerRmaItemHydBalanceValue    :" + actstockLedgerRmaItemHydBalanceValue);
		System.out.println("expstockLedgerRmaItemHydBalanceValue    :" + expstockLedgerRmaItemHydBalanceValue);

		System.out.println("actstockLedgerrmaItemHydAvgRate         :" + actstockLedgerrmaItemHydAvgRate);
		System.out.println("expstockLedgerrmaItemHydAvgRate         :" + expstockLedgerrmaItemHydAvgRate);
		
		click(report_CloseBtn);
		Thread.sleep(1000);
		click(stockLedgerHomePageFilterBtn);
		Thread.sleep(2000);
		click(stockLedgerFilterWarehouseField2);
		Thread.sleep(1000);
		stockLedgerFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(1000);
		stockLedgerFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(1000);
		enterText(stockLedgerFilterWarehouseField3, "SEC");
		Thread.sleep(1000);
		click(stockLedgerFilterOkBtn);
		Thread.sleep(1000);
		click(sl_OkBtn);
		Thread.sleep(1000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);
		
		
		
		String actstockLedgerRmaItemSecBalanceQty   = getText(stockLedgerRmaItemSecBalanceQty);
		String expstockLedgerRmaItemSecBalanceQty   = "1.0000";

		String actstockLedgerRmaItemSecBalanceValue = getText(stockLedgerRmaItemSecBalanceValue);
		String expstockLedgerRmaItemSecBalanceValue = "9.0458";

		String actstockLedgerrmaItemSecAvgRate      = getText(stockLedgerrmaItemSecAvgRate);
		String expstockLedgerrmaItemSecAvgRate      = "9.0458";
		
		System.out.println("=========actstockLedgerRmaItem  SEC WareHouse  Values============   : ");
		
		System.out.println("actstockLedgerRmaItemSecBalanceQty      :" + actstockLedgerRmaItemSecBalanceQty);
		System.out.println("expstockLedgerRmaItemSecBalanceQty      :" + expstockLedgerRmaItemSecBalanceQty);

		System.out.println("actstockLedgerRmaItemSecBalanceValue    :" + actstockLedgerRmaItemSecBalanceValue);
		System.out.println("expstockLedgerRmaItemSecBalanceValue    :" + expstockLedgerRmaItemSecBalanceValue);

		System.out.println("actstockLedgerrmaItemSecAvgRate         :" + actstockLedgerrmaItemSecAvgRate);
		System.out.println("expstockLedgerrmaItemSecAvgRate         :" + expstockLedgerrmaItemSecAvgRate);
	

		if (actstockLedgerRmaItemBalanceQty.equalsIgnoreCase(expstockLedgerRmaItemBalanceQty)
				&& actstockLedgerRmaItemBalanceValue.equalsIgnoreCase(expstockLedgerRmaItemBalanceValue)
				&& actstockLedgerrmaItemAvgRate.equalsIgnoreCase(expstockLedgerrmaItemAvgRate) && actstockLedgerRmaItemHydBalanceQty.equalsIgnoreCase(expstockLedgerRmaItemHydBalanceQty)
				&& actstockLedgerRmaItemHydBalanceValue.equalsIgnoreCase(expstockLedgerRmaItemHydBalanceValue) && actstockLedgerrmaItemHydAvgRate.equalsIgnoreCase(expstockLedgerrmaItemHydAvgRate)
				&& actstockLedgerRmaItemSecBalanceQty.equalsIgnoreCase(expstockLedgerRmaItemSecBalanceQty) && actstockLedgerRmaItemSecBalanceValue.equalsIgnoreCase(expstockLedgerRmaItemSecBalanceValue)
				&& actstockLedgerrmaItemSecAvgRate.equalsIgnoreCase(expstockLedgerrmaItemSecAvgRate)) {

			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return true;
		} else {
			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	@FindBy(xpath="//*[@id='61']/span")
	private static WebElement  financialsTransactionMenu; 

	@FindBy(xpath="//*[@id='2007']/span")
	private static WebElement  financialsTransactionsPurchaseMenu; 
	
	@FindBy(xpath="//*[@id='2014']/span")
	private static WebElement  financialsTransactionsSalesMenu; 
	
	
	
	@FindBy(xpath="//a[@id='2008']/span")
	private static WebElement  PurchasesVoucher; 
	
	@FindBy(xpath="//a[@id='2015']/span")
	private static WebElement  salesInvoiceVoucher;
	
	@FindBy(xpath="//a[@id='81']/span")
	public static WebElement  financialReports;
	
	@FindBy(xpath="//a[@id='500']/span")
	public static WebElement  financialReportsLedger;
	
	@FindBy(xpath="//a[@id='506']/span")
	public static WebElement  financialPurchaseReportMenu;
	
	@FindBy(xpath="//a[@id='507']/span")
	public static WebElement  financialPurchaseRegister;
	
	@FindBy(xpath="//a[@id='530']/span")
	public static WebElement  financialFinalAccountsMenu;
	
	@FindBy(xpath="//a[@id='558']/span")
	public static WebElement  financialReceivableAndPayableMenu;
	
	@FindBy(xpath="//a[@id='607']/span")
	public static WebElement  customerSummeryMenu;
	
	@FindBy(xpath="//a[@id='608']/span")
	public static WebElement  vendorSummeryMenu;
	
	@FindBy(xpath="//a[@id='544']/span")
	public static WebElement  vendorAgeingSummery;
	
	@FindBy(xpath="//a[@id='562']/span")
	public static WebElement  customerAgeingSummery;
	
	@FindBy(xpath="//a[@id='531']/span")
	public static WebElement  trialBalance;
	
	@FindBy(xpath="//div[text()='Financials']")
	private static WebElement financialsMenu;
	
	@FindBy(xpath="//div[@id='myNavbar']//span[@class='icon-new icon-font6']")
	private static WebElement  newBtn;
	
	@FindBy(xpath="//*[@id='id_transactionentry_previous']")
	private static WebElement  previousBtn;
	
	@FindBy(xpath = "//input[@id='id_header_1']")
	public static WebElement documentNumberTxt;
	
	
	
	@FindBy(xpath="//span[@class='icon-menu icon-font4']")
	private static WebElement  toggleBtn;
	
	@FindBy(xpath="//*[@id='id_transactionentry_copytoclipboard']")
	private static WebElement  copyToClipBoardOption;
	
	@FindBy(xpath="//*[@id='id_transactionentry_pastefromclipboard']")
	private static WebElement  pasteFromClipBoardOption;
	
	@FindBy(xpath="//input[@id='txtSearchMenu_MainLayout']")
	private static WebElement  homepageManuSearchField;
	
	@FindBy(xpath="//span[@id='2057']")
	private static WebElement  stockIn;
	
	@FindBy(xpath="//input[@id='id_header_3']")
	private static WebElement  purchasesAccounttxt;
	
	@FindBy(xpath="//input[@id='id_header_3']")
	private static WebElement  stockAccounttxt;
	
	@FindBy(xpath="//input[@id='id_header_3']")
	private static WebElement  salesAccounttxt;

	@FindBy(xpath="//input[@id='id_header_4']")
	private static WebElement  vendorAccounttxt;
	
	@FindBy(xpath="//input[@id='id_header_4']")
	private static WebElement  partyAccounttxt;
	
	@FindBy(xpath="//input[@id='id_header_4']")
	private static WebElement  customerAccounttxt;
	
	@FindBy(xpath="//input[@id='id_header_2']")
	private static WebElement  purchaseDatetxt;
	
	@FindBy(xpath="//input[@id='id_header_2']")
	private static WebElement  salesDatetxt;
	
	@FindBy(xpath="//input[@id='id_header_268435459']")
	private static WebElement  purchasesDepartmenttxt;
	
	@FindBy(xpath="//input[@id='id_header_268435459']")
	private static WebElement  stockOutDepartmenttxt;
	
	@FindBy(xpath="//input[@id='id_header_268435459']")
	private static WebElement  salesDepartmenttxt;
	
	
	@FindBy(xpath="//input[@id='id_body_38']")
	private static WebElement  enterRmatxt;
	
	@FindBy(xpath="//*[@id='id_transactionentry_save']")
	private static WebElement  voucherSaveBtn;
	
	@FindBy(xpath="//*[@id='id_transactionentry_close']")
	private static WebElement  new_CloseBtn;
	
	@FindBy(xpath="//div[@id='dvHomeTransClose']")  
	private static WebElement  voucherhomeCloseBtn;

	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[2]")
	private static WebElement  select1stRow_1stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[2]")
	private static WebElement  select2ndRow_1stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[2]")
	private static WebElement  select3rdRow_1stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[2]")
	private static WebElement  select4thRow_1stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[2]")
	private static WebElement  select5thRow_1stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[2]")
	private static WebElement  select6thRow_1stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[2]")
	private static WebElement  select7thRow_1stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[2]")
	private static WebElement  select8thRow_1stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[2]")
	private static WebElement  select9thRow_1stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[10]")
	private static WebElement  select1stRow_9thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[11]")
	private static WebElement  select1stRow_10thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[9]")
	private static WebElement  select1stRow_8thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[8]")
	private static WebElement  select1stRow_7thColumn;
	
	@FindBy(xpath="//i[@class='icon-pick icon-font5']")
	private static WebElement batchPickOnFIFOIcon;
	
	@FindBy(xpath="//label[@id='div_Ok']")
	private static WebElement batchOkIcon;
	
	@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_body']/tr[1]/td[3]")
	private static WebElement  batchPopUPRow1BaseQuantityTxt;
	
	@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_body']/tr[1]/td[8]")
	private static WebElement  batchPopUPRow1RateTxt;
	
	@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_body']/tr[2]/td[3]")
	private static WebElement  batchPopUPRow2BaseQuantityTxt;
	
	@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_body']/tr[2]/td[8]")
	private static WebElement  batchPopUPRow2RateTxt;
	
	@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_body']/tr[3]/td[3]")
	private static WebElement  batchPopUPRow3BaseQuantityTxt;
	
	@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_body']/tr[3]/td[8]")
	private static WebElement  batchPopUPRow3RateTxt;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[10]")
	private static WebElement  select2ndRow_9thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[11]")
	private static WebElement  select2ndRow_10thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[11]")
	private static WebElement  select3rdRow_10thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[12]")
	private static WebElement  select3rdRow_11thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[10]")
	private static WebElement  select3rdRow_9thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[11]")
	private static WebElement  select4thRow_10thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[12]")
	private static WebElement  select4thRow_11thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[12]")
	private static WebElement  select7thRow_11thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[09]")
	private static WebElement  select3rdRow_8thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[09]")
	private static WebElement  select4thRow_8thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[09]")
	private static WebElement  select5thRow_8thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[12]")
	private static WebElement  select5thRow_11thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[13]")
	private static WebElement  select5thRow_12thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[13]")
	private static WebElement  select8thRow_12thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[13]")
	private static WebElement  select9thRow_12thColumn;
	
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[09]")
	private static WebElement  select6thRow_8thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[12]")
	private static WebElement  select6thRow_11thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[13]")
	private static WebElement  select6thRow_12thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[9]")
	private static WebElement  select3rdRow_OverAllAvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[8]")
	private static WebElement  select3rdRow_AvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[8]")
	private static WebElement  select4thRow_AvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[9]")
	private static WebElement  select4thRow_OverAllAvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[8]")
	private static WebElement  select5thRow_AvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[9]")
	private static WebElement  select5thRow_OverAllAvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[8]")
	private static WebElement  select6thRow_AvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[10]")
	private static WebElement  select8thRow_AvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[10]")
	private static WebElement  select9thRow_AvgRate;
	
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[9]")
	private static WebElement  select6thRow_OverAllAvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[9]")
	private static WebElement  select8thRow_OverAllAvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[4]")
	private static WebElement  select8thRow_AvailaleStock;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[4]")
	private static WebElement  select1thRow_AvailaleStock;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[4]")
	private static WebElement  select4thRow_AvailaleStock;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[4]")
	private static WebElement  select6thRow_AvailaleStock;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[4]")
	private static WebElement  select7thRow_AvailaleStock;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[4]")
	private static WebElement  select9thRow_AvailaleStock;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[9]")
	private static WebElement select9thRow_OverAllAvgRate;

	// RMA POpup Window
	@FindBy(xpath = "//div[@id='Modal_Header']/div[1]")
	private static WebElement rmaScreenTitle;

	@FindBy(xpath = "//input[@id='txtSerialNo']")
	private static WebElement rmaSerialNumberTxtField;

	@FindBy(xpath = "//i[@id='btnSearch']")
	private static WebElement rmaSearchGridField;

	@FindBy(xpath = "//*[@id='RMASearchTable']/tbody/tr/td[2]")
	private static List<WebElement> rmaSerialNos;

	@FindBy(xpath = "//*[@id='RMASearchTable']/tbody/tr/td/input")
	private static List<WebElement> rmaSerialNoCheckBoxes;

	@FindBy(xpath = "//*[@id='RMA_Table_body']/tr/td[2]")
	private static List<WebElement> rmaSerialNoList;

	@FindBy(xpath = "//input[@id='chkSelectALL']")
	private static WebElement searchRmaSelectAllGridCheckBox;

	@FindBy(xpath = "//*[@id='SearchModal_Content']/div[3]/div/div[1]/label/i")
	private static WebElement searchRmaOkBtn;

	@FindBy(xpath = "//input[@id='txtQuantity']")
	private static WebElement rmaQuantityTxtField;

	@FindBy(xpath = "//div[@id='img_add']//span")
	private static WebElement rmaAddBtn;

	@FindBy(xpath = "//i[@class='icon-ok icon-font']")
	private static WebElement RMAPopupOkBtn;
	  
	///
	
	@FindBy(xpath = "//input[@id='id_body_536870916']")
	private static WebElement enterWarehousettxt;

	@FindBy(xpath = "//tbody[@id='id_body_536870916_table_body']/tr/td[2]")
	private static List<WebElement> svwareHouseListCount;

	@FindBy(xpath="//input[@id='id_body_23']")
	private static WebElement  enterItemttxt;
	
	@FindBy(xpath="//input[@id='id_body_33555468']")
	private static WebElement  enterAvailableStocktxt;
	
	
	@FindBy(xpath="//input[@id='id_body_26']")
	private static WebElement  enterQuantitytxt;
	
	@FindBy(xpath="//div[@id='idGlobalError']")
	public static WebElement validationConfirmationMessage;
	
	@FindBy(xpath="//input[@id='id_body_27']")
	private static WebElement  enterRatetxt;
	
	
	@FindBy(xpath="//input[@id='id_body_13']")
	private static WebElement  enterBatchtxt;
	
	@FindBy(xpath="//textarea[@id='id_body_16777279']")
	private static WebElement  enterBatCodetxt;
	
	@FindBy (xpath="//*[@id='TransactionNewBin']/div/div/div[2]/div[1]/div[7]/button")
	private static WebElement binSearchBtn;
	
	@FindBy (xpath="//div[@id='id_btnautoallocate']")
	private static WebElement binAutoAllocateBtn;
	
	@FindBy (xpath="//*[@id='id_transaction_bins_grid_col_1-0']")
	private static WebElement binPopup1stRow;
	
	@FindBy (xpath="//*[@id='TransactionNewBin']/div/div/div[3]/div[3]/div[3]")
	private static WebElement binPickBtn;
	
	@FindBy (xpath="//div[@id='id_bins_ok']//span[@class='icon-ok icon-font6']")
	private static WebElement binOkBtn;
	
	@FindBy (xpath="//input[@id='srch_bin']")
	private static WebElement binSearchTxt;
	
	@FindBy (xpath="//input[@id='id_body_84']")
	private static WebElement bin2Txt;
	
	@FindBy (xpath="//input[@id='id_body_85']")
	private static WebElement bin1Txt;
	
	public boolean checkSavePurchasesVoucher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		
		click(financialsMenu);
		click(financialsTransactionMenu);
		click(financialsTransactionsPurchaseMenu);
		click(PurchasesVoucher);
		click(newBtn);
		checkValidationMessage("Screen opened");
		click(purchaseDatetxt);
		purchaseDatetxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		purchaseDatetxt.sendKeys(Keys.BACK_SPACE);
		enterText(purchaseDatetxt, "01/09/2021");
		tab(purchaseDatetxt);
		Thread.sleep(2000);
		enterText(purchasesAccounttxt, "Cost of goods sold - Computers");
		tab(purchasesAccounttxt);
		Thread.sleep(2000);
		enterText(vendorAccounttxt, "Bank");
		Thread.sleep(2000);
		click(purchasesDepartmenttxt);
		enterText(purchasesDepartmenttxt, "DEPT1");
		tab(purchasesDepartmenttxt);
		Thread.sleep(3000);
		click(select1stRow_1stColumn);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "BR-BATCH");
		Thread.sleep(2000);
		tab(enterItemttxt);
		enterText(enterQuantitytxt, "10");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "1000");
		tab(enterRatetxt);
		Thread.sleep(2000);
		click(select1stRow_9thColumn);
		enterText(enterBatchtxt, "01092021PH");
		tab(enterBatchtxt);
		

		// 2nd Row

		click(select2ndRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "BR-BATCH");
		Thread.sleep(2000);
		tab(enterItemttxt);
		enterText(enterQuantitytxt, "10");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "2000");
		tab(enterRatetxt);
		click(select2ndRow_9thColumn);
		enterText(enterBatchtxt, "01092021PH");
		tab(enterBatchtxt);
		

		// 3rd Row

		click(select3rdRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		enterText(enterQuantitytxt, "10");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "1000");
		tab(enterRatetxt);
		click(select3rdRow_10thColumn);
		
		String actual3rdAvgRate 		= select3rdRow_AvgRate.getText();
		String exp3rdAvgRate 			= "8.0404";
		String actual3rdOverallAvgRate 	= select3rdRow_OverAllAvgRate.getText();
		String exp3rdoverallAvgRate 	= "8.2161";

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying + "  Value Expected  "
				+ expBinpopupIsDisplaying);

		click(binSearchBtn);
		Thread.sleep(3000);
		click(binAutoAllocateBtn);
		Thread.sleep(2000);
		click(binOkBtn);
		Thread.sleep(2000);
		

		// 4th Row

		click(select4thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		enterText(enterQuantitytxt, "10");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "2000");
		tab(enterRatetxt);
		click(select4thRow_10thColumn);
		
		String actual4throwAvgRate 			= select4thRow_AvgRate.getText();
		String exp4throwAvgRate		   	    = "8.4444";
		String actual4throwOverallAvgRate   = select4thRow_OverAllAvgRate.getText();
		String exp4throwoverallAvgRate      = "8.2161";

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying1 = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying1 = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying1 + "  Value Expected  "
				+ expBinpopupIsDisplaying1);

		click(binSearchBtn);
		Thread.sleep(2000);
		click(binAutoAllocateBtn);
		Thread.sleep(2000);
		click(binOkBtn);
		Thread.sleep(2000);
		

		// 5th Row

		click(select5thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);
		tab(enterItemttxt);
		enterText(enterQuantitytxt, "10");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "1000");
		tab(enterRatetxt);
		click(select5thRow_11thColumn);
		
		String actual5throwAvgRate 			= select5thRow_AvgRate.getText();
		String exp5throwAvgRate			    = "9.4927";
		String actual5throwOverallAvgRate   = select5thRow_OverAllAvgRate.getText();
		String exp5throwoverallAvgRate      = "9.6019";

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying + "  Value Expected  "
				+ expRMAPopupIsDisplaying);
		Thread.sleep(1000);
		enterText(rmaSerialNumberTxtField, "01092021PS,10");
		click(rmaAddBtn);
		Thread.sleep(2000);
		click(RMAPopupOkBtn);
		Thread.sleep(3000);
		tab(enterRmatxt);
		tab(enterBatCodetxt);

		// 6th Row
		Thread.sleep(1000);
		click(select6thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);
		tab(enterItemttxt);
		enterText(enterQuantitytxt, "10");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "2000");
		tab(enterRatetxt);
		click(select6thRow_11thColumn);
		
		String actual6throwAvgRate 				= select6thRow_AvgRate.getText();
		String exp6throwAvgRate 				= "9.0458";
		String actual6throwOverallAvgRate		= select6thRow_OverAllAvgRate.getText();
		String exp6throwoverallAvgRate			= "9.6019";

		System.out.println("actual3rdAvgRate       		 :" + actual3rdAvgRate);
		System.out.println("exp3rdAvgRate           	 :" + exp3rdAvgRate);

		System.out.println("actual3rdOverallAvgRate      :" + actual3rdOverallAvgRate);
		System.out.println("exp3rdoverallAvgRate         :" + exp3rdoverallAvgRate);

		System.out.println("actual4thAvgRate        	 :" + actual4throwAvgRate);
		System.out.println("exp4throwAvgRate          	 :" + exp4throwAvgRate);

		System.out.println("actual4throwOverallAvgRate   :" + actual4throwOverallAvgRate);
		System.out.println("exp4throwoverallAvgRate      :" + exp4throwoverallAvgRate);

		System.out.println("actual5throwAvgRate       	 :" + actual5throwAvgRate);
		System.out.println("exp5throwAvgRate             :" + exp5throwAvgRate);

		System.out.println("actual5throwOverallAvgRate   :" + actual5throwOverallAvgRate);
		System.out.println("exp5throwoverallAvgRate      	:" + exp5throwoverallAvgRate);

		System.out.println("actual6throwAvgRate      	 :" + actual6throwAvgRate);
		System.out.println("exp6throwAvgRate          	 :" + exp6throwAvgRate);

		System.out.println("actual6throwOverallAvgRate   :" + actual6throwOverallAvgRate);
		System.out.println("exp6throwoverallAvgRate      :" + exp6throwoverallAvgRate);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying1 = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying1 = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying1 + "  Value Expected  "
				+ expRMAPopupIsDisplaying1);
		enterText(rmaSerialNumberTxtField, "01092021PH9,10");
		click(rmaAddBtn);
		Thread.sleep(2000);
		click(RMAPopupOkBtn);
		Thread.sleep(2000);		
		tab(enterRmatxt);
		
		
		Thread.sleep(3000);
		click(voucherSaveBtn);

		HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) {
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : 35");
		expMessage.add("Saving in background.");

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage);

		if (actMessage.equals(expMessage) && actual3rdAvgRate.equalsIgnoreCase(exp3rdAvgRate)
				&& actual3rdOverallAvgRate.equalsIgnoreCase(exp3rdoverallAvgRate)
				&& actual4throwAvgRate.equalsIgnoreCase(exp4throwAvgRate)
				&& actual4throwOverallAvgRate.equalsIgnoreCase(exp4throwoverallAvgRate)
				&& actual5throwAvgRate.equalsIgnoreCase(exp5throwAvgRate)
				&& actual5throwOverallAvgRate.equalsIgnoreCase(exp5throwoverallAvgRate)
				&& actual6throwAvgRate.equalsIgnoreCase(exp6throwAvgRate)
				&& actual6throwOverallAvgRate.equalsIgnoreCase(exp6throwoverallAvgRate)) {
			
			click(new_CloseBtn);			
			Thread.sleep(1000);
			click(voucherhomeCloseBtn);
			return true;
			
		} else {
			click(new_CloseBtn);			
			Thread.sleep(1000);
			click(voucherhomeCloseBtn);
			return false;

		}

	}

	public boolean validteItemQueryStocksInBatchItemAfterSavingPurchaseVocher() throws InterruptedException {

		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(itemqueryMenu);
		click(itemTxt);
		Thread.sleep(2000);
		enterText(itemTxt, "BR-BATCH");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "25.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "1,201.6125";

		String actlstockValueTxt     = getText(stockValueTxt);
		String expstockValueTxt      = "30,040.3131";

		System.out.println("=====Stocks In Batch Items======");

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock           :" + actlCurrentStock);
		System.out.println("expCurrentStock            :" + expCurrentStock);

		System.out.println("actlavgStockRate           :" + actlavgStockRate);
		System.out.println("expavgStockRate            :" + expavgStockRate);

		System.out.println("actlstockValueTxt          :" + actlstockValueTxt);
		System.out.println("expstockValueTxt           :" + expstockValueTxt);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)) {
			return true;
		} else {
			return false;
		}

	}

	public boolean validteItemQueryStocksInFifoItemAfterSavingPurchaseVocher() throws InterruptedException {

		click(itemTxt);
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.BACK_SPACE);
		enterText(itemTxt, "FIFO-BIN");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "25.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "1,201.6432";

		String actlstockValueTxt     = getText(stockValueTxt);
		String expstockValueTxt      = "30,041.0807";

		System.out.println("=====Stocks In FIFO Items======");

		System.out.println("actlOpenStockQuantity :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity  :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock      :" + actlCurrentStock);
		System.out.println("expCurrentStock       :" + expCurrentStock);

		System.out.println("actlavgStockRate      :" + actlavgStockRate);
		System.out.println("expavgStockRate       :" + expavgStockRate);

		System.out.println("actlstockValueTxt     :" + actlstockValueTxt);
		System.out.println("expstockValueTxt      :" + expstockValueTxt);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)) {
			return true;
		} else {
			return false;
		}

	}

	public boolean validteItemQueryStocksInWAItemAfterSavingPurchaseVocher() throws InterruptedException {
		
		Thread.sleep(2000);
		click(itemTxt);
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.BACK_SPACE);
		enterText(itemTxt, "WA-RMA");
		Thread.sleep(4000);
		click(getStockBtn);
		Thread.sleep(4000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "25.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "1,201.9204";

		String actlstockValueTxt     = getText(stockValueTxt);
		String expstockValueTxt      = "30,048.0097";

		ScrollToElement(balanceByWarehouseBtn);
		click(balanceByWarehouseBtn);

		Thread.sleep(4000);

		int count = row1BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow1 = new ArrayList<String>();

		for (int i = 0; i < count; i++) {
			String data = row1BalanceByWareHouseList.get(i).getText();
			warehouseRow1.add(data);
		}

		String actwarehouseRow1List = warehouseRow1.toString();
		String expWarehouseRow1List = "[1, HYD, 12.000, 0.000, 12.000, 834.9155, 10,018.9854]";

		int count1 = row2BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow2 = new ArrayList<String>();

		for (int i = 0; i < count1; i++) {
			String data = row2BalanceByWareHouseList.get(i).getText();
			warehouseRow2.add(data);
		}

		String actwarehouseRow2List = warehouseRow2.toString();
		String expWarehouseRow2List = "[2, SEC, 11.000, 0.000, 11.000, 1,819.0042, 20,009.0458]";

		int count2 = row3BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow3 = new ArrayList<String>();

		for (int i = 0; i < count2; i++) {
			String data = row3BalanceByWareHouseList.get(i).getText();
			warehouseRow3.add(data);
		}

		String actwarehouseRow3List = warehouseRow3.toString();
		String expWarehouseRow3List = "[3, TEST, 2.000, 0.000, 2.000, 11.3062, 22.6125]";

		System.out.println("=====Stocks In WA Items======");

		System.out.println("actlOpenStockQuantity :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity  :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock      :" + actlCurrentStock);
		System.out.println("expCurrentStock       :" + expCurrentStock);

		System.out.println("actlavgStockRate      :" + actlavgStockRate);
		System.out.println("expavgStockRate       :" + expavgStockRate);

		System.out.println("actlstockValueTxt     :" + actlstockValueTxt);
		System.out.println("expstockValueTxt      :" + expstockValueTxt);

		System.out.println("===========Balance By Warehouse=======");

		System.out.println(actwarehouseRow1List);
		System.out.println(expWarehouseRow1List);

		System.out.println(actwarehouseRow2List);
		System.out.println(expWarehouseRow2List);

		System.out.println(actwarehouseRow3List);
		System.out.println(expWarehouseRow3List);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)
				&& actwarehouseRow1List.equalsIgnoreCase(expWarehouseRow1List)
				&& actwarehouseRow2List.equalsIgnoreCase(expWarehouseRow2List)
				&& actwarehouseRow3List.equalsIgnoreCase(expWarehouseRow3List)) {
			return true;
		} else {
			return false;
		}

	}  

	public boolean checkStockLedgerReportFortBrBatchItemAfterSavingPurchaseVocher()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryStockLedgerMenu);
		Thread.sleep(3000);

		int rowcount = stockLedgerHometableRowCount.size();

		System.out.println(rowcount);

		for (int i = 0; i < rowcount; i++) {
			String actName = ledgerHometableItemNamesList.get(i).getText();

			System.out.println(actName);

			if (actName.equalsIgnoreCase("BR-BATCH")) {
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);

		String actStockLedgerBrBatchBalanceQty   = getText(stockLedgerBrBatchAfterSavingpurchaseBalanceQty);
		String expStockLedgerBrBatchBalanceQty   = "25.0000";

		String actStockLedgerBrBatchBalanceValue = getText(stockLedgerBrBatchBalanceValueAfterSavingpurchase);
		String expStockLedgerBrBatchBalanceValue = "30,040.3131";

		String actStockLedgerBrBatchAvgRate      = getText(stockLedgerBrBatchAvgRateAfterSavingpurchase);
		String expStockLedgerBrBatchAvgRate      = "1,201.6125";
		

		System.out.println("actStockLedgerBrBatchBalanceQty        :" + actStockLedgerBrBatchBalanceQty);
		System.out.println("expStockLedgerBrBatchBalanceQty        :" + expStockLedgerBrBatchBalanceQty);

		System.out.println("actStockLedgerBrBatchBalanceValue      :" + actStockLedgerBrBatchBalanceValue);
		System.out.println("expStockLedgerBrBatchBalanceValue      :" + expStockLedgerBrBatchBalanceValue);

		System.out.println("actStockLedgerBrBatchAvgRate      	   :" + actStockLedgerBrBatchAvgRate);
		System.out.println("expStockLedgerBrBatchAvgRate    	   :" + expStockLedgerBrBatchAvgRate);

		if (actStockLedgerBrBatchBalanceQty.equalsIgnoreCase(expStockLedgerBrBatchBalanceQty)
				&& actStockLedgerBrBatchBalanceValue.equalsIgnoreCase(expStockLedgerBrBatchBalanceValue)
				&& actStockLedgerBrBatchAvgRate.equalsIgnoreCase(expStockLedgerBrBatchAvgRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}

	public boolean checkStockLedgerReportForFifoItemAfterSavingPurchaseVocher()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == false) {
			click(fifoItemCheckBox);
		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == true) {
			click(rmaItemCheckBox);

		}

		
		dropDown(sl_DateOptionDropdown, "1");	
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);

		String actstockLedgerFifoItemBalanceQty 	= getText(stockLedgerFifoItemBalanceQtyAfterSavingPurchase);
		String expstockLedgerFifoItemBalanceQty 	= "25.0000";

		String actstockLedgerFifoItemBalanceValue 	= getText(stockLedgerFifoItemBalanceValueAfterSavingPurchase);
		String expstockLedgerFifoItemBalanceValue   = "30,041.0807";

		String actstockLedgerFifoItemAvgRate		= getText(stockLedgerFifoItemAvgRateAfterSavingPurchase);
		String expstockLedgerFifoItemAvgRate		= "1,201.6432";

		System.out.println("actstockLedgerFifoItemBalanceQty      	:" + actstockLedgerFifoItemBalanceQty);
		System.out.println("expstockLedgerFifoItemBalanceQty      	:" + expstockLedgerFifoItemBalanceQty);

		System.out.println("actstockLedgerFifoItemBalanceValue      :" + actstockLedgerFifoItemBalanceValue);
		System.out.println("expstockLedgerFifoItemBalanceValue      :" + expstockLedgerFifoItemBalanceValue);

		System.out.println("actstockLedgerFifoItemAvgRate     		:" + actstockLedgerFifoItemAvgRate);
		System.out.println("expstockLedgerFifoItemAvgRate      		:" + expstockLedgerFifoItemAvgRate);

		if (actstockLedgerFifoItemBalanceQty.equalsIgnoreCase(expstockLedgerFifoItemBalanceQty)
				&& actstockLedgerFifoItemBalanceValue.equalsIgnoreCase(expstockLedgerFifoItemBalanceValue)
				&& actstockLedgerFifoItemAvgRate.equalsIgnoreCase(expstockLedgerFifoItemAvgRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}
	
	@FindBy(xpath = "(//tr[@id='trRender_4'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemHydBalanceQtyAfterSavingPurchaseVocher;
	
	@FindBy(xpath = "(//tr[@id='trRender_4'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemHydBalanceValueAfterSavingPurchaseVocher;
	
	@FindBy(xpath = "(//tr[@id='trRender_4'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemHydAvgRateAfterSavingPurchaseVocher;
	
	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemSecBalanceQtyAfterSavingPurchaseVocher;
	
	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemSecBalanceValueAfterSavingPurchaseVocher;
	
	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemSecAvgRateAfterSavingPurchaseVocher;

	public boolean checkStockLedgerReportForRmaItemAfterSavingPurchaseVocher()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == true) {
			click(fifoItemCheckBox);

		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == false) {
			click(rmaItemCheckBox);

		}

	
		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(3000);

		String actstockLedgerRmaItemBalanceQty 		= getText(stockLedgerRmaItemBalanceQtyAftersavingPurchase);
		String expstockLedgerRmaItemBalanceQty 		= "25.0000";

		String actstockLedgerRmaItemBalanceValue 	= getText(stockLedgerRmaItemBalanceValueAfterSavingPurchase);
		String expstockLedgerRmaItemBalanceValue 	= "30,048.0097";

		String actstockLedgerrmaItemAvgRate 		= getText(stockLedgerrmaItemAvgRateAfterSavingPurchase);
		String expstockLedgerrmaItemAvgRate 		= "1,201.9204";

		System.out.println("actstockLedgerRmaItemBalanceQty      	:" + actstockLedgerRmaItemBalanceQty);
		System.out.println("expstockLedgerRmaItemBalanceQty      	:" + expstockLedgerRmaItemBalanceQty);

		System.out.println("actstockLedgerRmaItemBalanceValue       :" + actstockLedgerRmaItemBalanceValue);
		System.out.println("expstockLedgerRmaItemBalanceValue       :" + expstockLedgerRmaItemBalanceValue);

		System.out.println("actstockLedgerrmaItemAvgRate      		:" + actstockLedgerrmaItemAvgRate);
		System.out.println("expstockLedgerrmaItemAvgRate      		:" + expstockLedgerrmaItemAvgRate);
		
		click(stockLedgerFilterBtn);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

		boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
		boolean expFilterpopupIsDisplaying = true;

		System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
				+ expFilterpopupIsDisplaying);
		
		click(stockLedgerFilterWarehouseField);
		enterText(stockLedgerFilterWarehouseField, "HYD");
		Thread.sleep(2000);
		click(stockLedgerFilterOkBtn);
		Thread.sleep(1000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);
		
		String actstockLedgerRmaItemHydBalanceQtyAfterSavingPurchaseVocher 		= getText(stockLedgerRmaItemHydBalanceQtyAfterSavingPurchaseVocher);
		String expstockLedgerRmaItemHydBalanceQtyAfterSavingPurchaseVocher 		= "12.0000";

		String actstockLedgerRmaItemHydBalanceValueAfterSavingPurchaseVocher 	= getText(stockLedgerRmaItemHydBalanceValueAfterSavingPurchaseVocher);
		String expstockLedgerRmaItemHydBalanceValueAfterSavingPurchaseVocher 	= "10,018.9854";

		String actstockLedgerrmaItemHydAvgRateAfterSavingPurchaseVocher 		= getText(stockLedgerrmaItemHydAvgRateAfterSavingPurchaseVocher);
		String expstockLedgerrmaItemHydAvgRateAfterSavingPurchaseVocher 		= "834.9155";
		
		System.out.println("=========actstockLedgerRmaItem  HydWareHouse  Values============   : ");
		
		System.out.println("actstockLedgerRmaItemHydBalanceQtyAfterSavingPurchaseVocher      :" + actstockLedgerRmaItemHydBalanceQtyAfterSavingPurchaseVocher);
		System.out.println("expstockLedgerRmaItemHydBalanceQtyAfterSavingPurchaseVocher      :" + expstockLedgerRmaItemHydBalanceQtyAfterSavingPurchaseVocher);
		System.out.println("actstockLedgerRmaItemHydBalanceValueAfterSavingPurchaseVocher    :" + actstockLedgerRmaItemHydBalanceValueAfterSavingPurchaseVocher);
		System.out.println("expstockLedgerRmaItemHydBalanceValueAfterSavingPurchaseVocher    :" + expstockLedgerRmaItemHydBalanceValueAfterSavingPurchaseVocher);

		System.out.println("actstockLedgerrmaItemHydAvgRateAfterSavingPurchaseVocher         :" + actstockLedgerrmaItemHydAvgRateAfterSavingPurchaseVocher);
		System.out.println("expstockLedgerrmaItemHydAvgRateAfterSavingPurchaseVocher         :" + expstockLedgerrmaItemHydAvgRateAfterSavingPurchaseVocher);		
		
		click(report_CloseBtn);
		Thread.sleep(1000);
		click(stockLedgerHomePageFilterBtn);
		Thread.sleep(2000);
		click(stockLedgerFilterWarehouseField2);
		Thread.sleep(1000);
		stockLedgerFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(1000);
		stockLedgerFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(1000);
		enterText(stockLedgerFilterWarehouseField3, "SEC");
		Thread.sleep(1000);
		click(stockLedgerFilterOkBtn);
		Thread.sleep(1000);
		click(sl_OkBtn);
		Thread.sleep(1000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);
		
		
		
		String actstockLedgerRmaItemSecBalanceQtyAfterSavingPurchaseVocher 		= getText(stockLedgerRmaItemSecBalanceQtyAfterSavingPurchaseVocher);
		String expstockLedgerRmaItemSecBalanceQtyAfterSavingPurchaseVocher 		= "11.0000";

		String actstockLedgerRmaItemSecBalanceValueAfterSavingPurchaseVocher 	= getText(stockLedgerRmaItemSecBalanceValueAfterSavingPurchaseVocher);
		String expstockLedgerRmaItemSecBalanceValueAfterSavingPurchaseVocher 	= "20,009.0458";

		String actstockLedgerrmaItemSecAvgRateAfterSavingPurchaseVocher 		= getText(stockLedgerrmaItemSecAvgRateAfterSavingPurchaseVocher);
		String expstockLedgerrmaItemSecAvgRateAfterSavingPurchaseVocher 		= "1,819.0042";
		
		System.out.println("=========actstockLedgerRmaItem  SEC WareHouse  Values============   : ");
		
		System.out.println("actstockLedgerRmaItemSecBalanceQtyAfterSavingPurchaseVocher      :" + actstockLedgerRmaItemSecBalanceQtyAfterSavingPurchaseVocher);
		System.out.println("expstockLedgerRmaItemSecBalanceQtyAfterSavingPurchaseVocher      :" + expstockLedgerRmaItemSecBalanceQtyAfterSavingPurchaseVocher);

		System.out.println("actstockLedgerRmaItemSecBalanceValueAfterSavingPurchaseVocher    :" + actstockLedgerRmaItemSecBalanceValueAfterSavingPurchaseVocher);
		System.out.println("expstockLedgerRmaItemSecBalanceValueAfterSavingPurchaseVocher    :" + expstockLedgerRmaItemSecBalanceValueAfterSavingPurchaseVocher);

		System.out.println("actstockLedgerrmaItemSecAvgRateAfterSavingPurchaseVocher         :" + actstockLedgerrmaItemSecAvgRateAfterSavingPurchaseVocher);
		System.out.println("expstockLedgerrmaItemSecAvgRateAfterSavingPurchaseVocher         :" + expstockLedgerrmaItemSecAvgRateAfterSavingPurchaseVocher);
	

		
		

		if (actstockLedgerRmaItemBalanceQty.equalsIgnoreCase(expstockLedgerRmaItemBalanceQty)
				&& actstockLedgerRmaItemBalanceValue.equalsIgnoreCase(expstockLedgerRmaItemBalanceValue)
				&& actstockLedgerrmaItemAvgRate.equalsIgnoreCase(expstockLedgerrmaItemAvgRate)
				&& actstockLedgerRmaItemHydBalanceQtyAfterSavingPurchaseVocher.equalsIgnoreCase(expstockLedgerRmaItemHydBalanceQtyAfterSavingPurchaseVocher)
				&& actstockLedgerRmaItemHydBalanceValueAfterSavingPurchaseVocher.equalsIgnoreCase(expstockLedgerRmaItemHydBalanceValueAfterSavingPurchaseVocher)
				&& actstockLedgerrmaItemHydAvgRateAfterSavingPurchaseVocher.equalsIgnoreCase(expstockLedgerrmaItemHydAvgRateAfterSavingPurchaseVocher)
				&& actstockLedgerRmaItemSecBalanceQtyAfterSavingPurchaseVocher.equalsIgnoreCase(expstockLedgerRmaItemSecBalanceQtyAfterSavingPurchaseVocher)
				&& actstockLedgerRmaItemSecBalanceValueAfterSavingPurchaseVocher.equalsIgnoreCase(expstockLedgerRmaItemSecBalanceValueAfterSavingPurchaseVocher)
				&& actstockLedgerrmaItemSecAvgRateAfterSavingPurchaseVocher.equalsIgnoreCase(expstockLedgerrmaItemSecAvgRateAfterSavingPurchaseVocher)) {

			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return true;
		} else {
			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	@FindBy(xpath="//span[@class='hidden-xs']")
	public static WebElement userNameDisplay;
	
	@FindBy(xpath="//*[@id='companyLogo']")
	public static WebElement companyLogo;

	@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
	public static WebElement companyName;

	@FindBy(xpath="//*[@id='userprofile']/li/span[2]")
	public static WebElement logoutOption;
	
	 @FindBy(id="ddlCompany")
	 public static WebElement companyDropDownList;

	
	public boolean checkLoginToInvUserAndValidteItemQueryStocksInBatchItem() throws InterruptedException
	{
		Thread.sleep(3000);
		click(userNameDisplay);
		Thread.sleep(1000);
		click(logoutOption);		
		Thread.sleep(3000);
		
		String unamelt="inv";

		String pawslt="su";
		
		LoginPage.enterUserName(unamelt);	     
        LoginPage.enterPassword(pawslt);
        
		
		 Select company = new Select(companyDropDownList);
		 company.selectByValue("108");
		 Thread.sleep(2000);
		 
  
        LoginPage.clickOnSignInBtn();  
        
        Thread.sleep(3000);
        
        System.out.println("Signing");
        
        // Login User Name is Verify
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		String actUserInfo                      = userNameDisplay.getText();
		String expUserInfo                      = "INV";	

		System.out.println("Login With Valid User and Check User Name Display Value Actual : ." + actUserInfo + ". Value Expected : " + expUserInfo);
		
		
		// Login Company Name is Verify
		Thread.sleep(1000);
		click(companyLogo);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
		String actGetLoginCompanyNameInformation = companyName.getText();
		String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
		String expGetLoginCompanyName            = "Maintain Quantity a";   
		companyLogo.click();
		
		System.out.println("Company Name Display Value Actual                : ." + actGetLoginCompanyName +  ". Value Expected : " + expGetLoginCompanyName);
		
		
		Thread.sleep(2000);		
		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(itemqueryMenu);
		click(itemTxt);
		Thread.sleep(2000);
		enterText(itemTxt, "BR-BATCH");
		Thread.sleep(2000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity 	= getText(openStockQuantityTxt);
		String expOpenStockQuantity 	= "20.000";

		String actlCurrentStock 		= getText(currentStockTxt);
		String expCurrentStock 			= "23.000";

		String actlavgStockRate 		= getText(avgStockRateTxt);
		String expavgStockRate 			= "1,305.3953";

		String actlstockValueTxt 		= getText(stockValueTxt);
		String expstockValueTxt 		= "30,024.0909";

		System.out.println("=====Stocks In Batch Items======");

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock      	   :" + actlCurrentStock);
		System.out.println("expCurrentStock       	   :" + expCurrentStock);

		System.out.println("actlavgStockRate     	   :" + actlavgStockRate);
		System.out.println("expavgStockRate      	   :" + expavgStockRate);

		System.out.println("actlstockValueTxt    	   :" + actlstockValueTxt);
		System.out.println("expstockValueTxt       	   :" + expstockValueTxt);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)) {
			return true;
		} else {
			return false;
		}
		
		}
	
	
	public boolean validteItemQueryStocksInFifoItemAfterLoginToInvUser() throws InterruptedException {

		click(itemTxt);
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.BACK_SPACE);
		enterText(itemTxt, "FIFO-BIN");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock 	 = getText(currentStockTxt);
		String expCurrentStock 		 = "23.000";

		String actlavgStockRate 	 = getText(avgStockRateTxt);
		String expavgStockRate 		 = "1,305.4317";

		String actlstockValueTxt 	 = getText(stockValueTxt);
		String expstockValueTxt 	 = "30,024.9292";

		System.out.println("=====Stocks In FIFO Items======");

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock      	   :" + actlCurrentStock);
		System.out.println("expCurrentStock            :" + expCurrentStock);

		System.out.println("actlavgStockRate           :" + actlavgStockRate);
		System.out.println("expavgStockRate            :" + expavgStockRate);

		System.out.println("actlstockValueTxt          :" + actlstockValueTxt);
		System.out.println("expstockValueTxt           :" + expstockValueTxt);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)) {
			return true;
		} else {
			return false;
		}

	}
	
	
	public boolean validteItemQueryStocksInWAItemAfterLoginToInvUser() throws InterruptedException {

		click(itemTxt);
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.BACK_SPACE);
		enterText(itemTxt, "WA-RMA");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock 	 = getText(currentStockTxt);
		String expCurrentStock       = "23.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "1,305.5580";

		String actlstockValueTxt     = getText(stockValueTxt);
		String expstockValueTxt      = "30,027.8344";

		ScrollToElement(balanceByWarehouseBtn);
		click(balanceByWarehouseBtn);
		Thread.sleep(3000);

		int count = row1BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow1 = new ArrayList<String>();

		for (int i = 0; i < count; i++) {
			String data = row1BalanceByWareHouseList.get(i).getText();
			warehouseRow1.add(data);
		}

		String actwarehouseRow1List = warehouseRow1.toString();
		String expWarehouseRow1List = "[1, HYD, 12.000, 0.000, 12.000, 834.9155, 10,018.9854]";

		int count1 = row2BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow2 = new ArrayList<String>();

		for (int i = 0; i < count1; i++) {
			String data = row2BalanceByWareHouseList.get(i).getText();
			warehouseRow2.add(data);
		}

		String actwarehouseRow2List = warehouseRow2.toString();
		String expWarehouseRow2List = "[2, SEC, 11.000, 0.000, 11.000, 1,819.0042, 20,009.0458]";

		System.out.println("=====Stocks In WA Items======");

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock           :" + actlCurrentStock);
		System.out.println("expCurrentStock            :" + expCurrentStock);

		System.out.println("actlavgStockRate           :" + actlavgStockRate);
		System.out.println("expavgStockRate            :" + expavgStockRate);

		System.out.println("actlstockValueTxt          :" + actlstockValueTxt);
		System.out.println("expstockValueTxt           :" + expstockValueTxt);

		System.out.println("===========Balance By Warehouse=======");

		System.out.println(actwarehouseRow1List);
		System.out.println(expWarehouseRow1List);

		System.out.println(actwarehouseRow2List);
		System.out.println(expWarehouseRow2List);

		
		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)
				&& actwarehouseRow1List.equalsIgnoreCase(expWarehouseRow1List)
				&& actwarehouseRow2List.equalsIgnoreCase(expWarehouseRow2List)) {
			return true;
		} else {
			return false;
		}

	}
	
	
	public boolean checkStockLedgerReportForBrBatchItemAfterLoginToInvUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryStockLedgerMenu);
		Thread.sleep(3000);

		int rowcount = stockLedgerHometableRowCount.size();

		System.out.println(rowcount);

		for (int i = 0; i < rowcount; i++) {
			String actName = ledgerHometableItemNamesList.get(i).getText();

			System.out.println(actName);

			if (actName.equalsIgnoreCase("BR-BATCH")) {
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}

		
		dropDown(sl_DateOptionDropdown, "1");	
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(3000);

		String actStockLedgerBrBatchBalanceQty   = getText(stockLedgerBrBatchBalanceQtyAfterLoginToInvUser);
		String expStockLedgerBrBatchBalanceQty   = "23.0000";

		String actStockLedgerBrBatchBalanceValue = getText(stockLedgerBrBatchBalanceValueAfterLoginToInvUser);
		String expStockLedgerBrBatchBalanceValue = "30,024.0909";

		String actStockLedgerBrBatchAvgRate      = getText(stockLedgerBrBatchAvgRateAfterLoginToInvUser);
		String expStockLedgerBrBatchAvgRate      = "1,305.3953";

		System.out.println("actStockLedgerBrBatchBalanceQty      :" + actStockLedgerBrBatchBalanceQty);
		System.out.println("expStockLedgerBrBatchBalanceQty      :" + expStockLedgerBrBatchBalanceQty);

		System.out.println("actStockLedgerBrBatchBalanceValue    :" + actStockLedgerBrBatchBalanceValue);
		System.out.println("expStockLedgerBrBatchBalanceValue    :" + expStockLedgerBrBatchBalanceValue);

		System.out.println("actStockLedgerBrBatchAvgRate         :" + actStockLedgerBrBatchAvgRate);
		System.out.println("expStockLedgerBrBatchAvgRate         :" + expStockLedgerBrBatchAvgRate);

		if (actStockLedgerBrBatchBalanceQty.equalsIgnoreCase(expStockLedgerBrBatchBalanceQty)
				&& actStockLedgerBrBatchBalanceValue.equalsIgnoreCase(expStockLedgerBrBatchBalanceValue)
				&& actStockLedgerBrBatchAvgRate.equalsIgnoreCase(expStockLedgerBrBatchAvgRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}
	
	
	public boolean checkStockLedgerReportForFifoItemAfterLoginToInvUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == false) {
			click(fifoItemCheckBox);
		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == true) {
			click(rmaItemCheckBox);

		}

		
		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(3000);

		String actstockLedgerFifoItemBalanceQty    = getText(stockLedgerFifoItemBalanceQtyAfterLoginToInvUser);
		String expstockLedgerFifoItemBalanceQty    = "23.0000";

		String actstockLedgerFifoItemBalancevalu   = getText(stockLedgerFifoItemBalancevaluAfterLoginToInvUser);
		String expstockLedgerFifoItemBalancevalu   = "30,024.9292";

		String actstockLedgerFifoItemAvgRate       = getText(stockLedgerFifoItemAvgRateAfterLoginToInvUser);
		String expstockLedgerFifoItemAvgRate       = "1,305.4317";

		System.out.println("actstockLedgerFifoItemBalanceQty      :" + actstockLedgerFifoItemBalanceQty);
		System.out.println("expstockLedgerFifoItemBalanceQty      :" + expstockLedgerFifoItemBalanceQty);

		System.out.println("actstockLedgerFifoItemBalancevalu     :" + actstockLedgerFifoItemBalancevalu);
		System.out.println("expstockLedgerFifoItemBalancevalu     :" + expstockLedgerFifoItemBalancevalu);

		System.out.println("actstockLedgerFifoItemAvgRate         :" + actstockLedgerFifoItemAvgRate);
		System.out.println("expstockLedgerFifoItemAvgRate         :" + expstockLedgerFifoItemAvgRate);

		if (actstockLedgerFifoItemBalanceQty.equalsIgnoreCase(expstockLedgerFifoItemBalanceQty)
				&& actstockLedgerFifoItemBalancevalu.equalsIgnoreCase(expstockLedgerFifoItemBalancevalu)
				&& actstockLedgerFifoItemAvgRate.equalsIgnoreCase(expstockLedgerFifoItemAvgRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}
	
	@FindBy(xpath = "(//tr[@id='trRender_4'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemHydBalanceQtyAfterLoginToInvUser;
	
	@FindBy(xpath = "(//tr[@id='trRender_4'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemHydBalanceValueAfterLoginToInvUser;
	
	@FindBy(xpath = "(//tr[@id='trRender_4'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemHydAvgRateAfterLoginToInvUser;
	
	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemSecBalanceQtyAfterLoginToInvUser;
	
	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemSecBalanceValueAfterLoginToInvUser;
	
	@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemSecAvgRateAfterLoginToInvUser;
	
	
	public boolean checkStockLedgerReportForRmaItemAfterLoginToInvUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == true) {
			click(fifoItemCheckBox);

		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == false) {
			click(rmaItemCheckBox);

		}

		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(3000);

		String actstockLedgerRmaItemBalanceQty   = getText(stockLedgerRmaItemBalanceQtyAfterLoginToInvUser);
		String expstockLedgerRmaItemBalanceQty   = "23.0000";

		String actstockLedgerRmaItemBalanceValue = getText(stockLedgerRmaItemBalanceValueAfterLoginToInvUser);
		String expstockLedgerRmaItemBalanceValue = "30,027.8344";

		String actstockLedgerrmaItemAvgRate      = getText(stockLedgerrmaItemAvgRateAfterLogintoInvUser);
		String expstockLedgerrmaItemAvgRate      = "1,305.5580";

		System.out.println("actstockLedgerRmaItemBalanceQty      :" + actstockLedgerRmaItemBalanceQty);
		System.out.println("expstockLedgerRmaItemBalanceQty      :" + expstockLedgerRmaItemBalanceQty);

		System.out.println("actstockLedgerRmaItemBalanceValue    :" + actstockLedgerRmaItemBalanceValue);
		System.out.println("expstockLedgerRmaItemBalanceValue    :" + expstockLedgerRmaItemBalanceValue);

		System.out.println("actstockLedgerrmaItemAvgRate         :" + actstockLedgerrmaItemAvgRate);
		System.out.println("expstockLedgerrmaItemAvgRate         :" + expstockLedgerrmaItemAvgRate);
		
		click(stockLedgerFilterBtn);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

		boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
		boolean expFilterpopupIsDisplaying = true;

		System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
				+ expFilterpopupIsDisplaying);
		
		click(stockLedgerFilterWarehouseField);
		enterText(stockLedgerFilterWarehouseField, "HYD");
		Thread.sleep(2000);
		click(stockLedgerFilterOkBtn);
		Thread.sleep(1000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);
		
		String actstockLedgerRmaItemHydBalanceQtyAfterLoginToInvUser   = getText(stockLedgerRmaItemHydBalanceQtyAfterLoginToInvUser);
		String expstockLedgerRmaItemHydBalanceQtyAfterLoginToInvUser   = "12.0000";

		String actstockLedgerRmaItemHydBalanceValueAfterLoginToInvUser = getText(stockLedgerRmaItemHydBalanceValueAfterLoginToInvUser);
		String expstockLedgerRmaItemHydBalanceValueAfterLoginToInvUser = "10,018.9854";

		String actstockLedgerrmaItemHydAvgRateAfterLoginToInvUser      = getText(stockLedgerrmaItemHydAvgRateAfterLoginToInvUser);
		String expstockLedgerrmaItemHydAvgRateAfterLoginToInvUser      = "834.9155";
		
		System.out.println("=========actstockLedgerRmaItem  HydWareHouse  Values============   : ");
		
		System.out.println("actstockLedgerRmaItemHydBalanceQtyAfterLoginToInvUser      :" + actstockLedgerRmaItemHydBalanceQtyAfterLoginToInvUser);
		System.out.println("expstockLedgerRmaItemHydBalanceQtyAfterLoginToInvUser      :" + expstockLedgerRmaItemHydBalanceQtyAfterLoginToInvUser);
		System.out.println("actstockLedgerRmaItemHydBalanceValueAfterLoginToInvUser    :" + actstockLedgerRmaItemHydBalanceValueAfterLoginToInvUser);
		System.out.println("expstockLedgerRmaItemHydBalanceValueAfterLoginToInvUser    :" + expstockLedgerRmaItemHydBalanceValueAfterLoginToInvUser);

		System.out.println("actstockLedgerrmaItemHydAvgRateAfterLoginToInvUser         :" + actstockLedgerrmaItemHydAvgRateAfterLoginToInvUser);
		System.out.println("expstockLedgerrmaItemHydAvgRateAfterLoginToInvUser         :" + expstockLedgerrmaItemHydAvgRateAfterLoginToInvUser);		
		
		click(report_CloseBtn);
		Thread.sleep(1000);
		click(stockLedgerHomePageFilterBtn);
		Thread.sleep(2000);
		click(stockLedgerFilterWarehouseField2);
		Thread.sleep(1000);
		stockLedgerFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(1000);
		stockLedgerFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(1000);
		enterText(stockLedgerFilterWarehouseField3, "SEC");
		Thread.sleep(1000);
		click(stockLedgerFilterOkBtn);
		Thread.sleep(1000);
		click(sl_OkBtn);
		Thread.sleep(1000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);
		
		
		
		String actstockLedgerRmaItemSecBalanceQtyAfterLoginToInvUser   = getText(stockLedgerRmaItemSecBalanceQtyAfterLoginToInvUser);
		String expstockLedgerRmaItemSecBalanceQtyAfterLoginToInvUser   = "11.0000";

		String actstockLedgerRmaItemSecBalanceValueAfterLoginToInvUser = getText(stockLedgerRmaItemSecBalanceValueAfterLoginToInvUser);
		String expstockLedgerRmaItemSecBalanceValueAfterLoginToInvUser = "20,009.0458";

		String actstockLedgerrmaItemSecAvgRateAfterLoginToInvUser      = getText(stockLedgerrmaItemSecAvgRateAfterLoginToInvUser);
		String expstockLedgerrmaItemSecAvgRateAfterLoginToInvUser      = "1,819.0042";
		
		System.out.println("=========actstockLedgerRmaItem  SEC WareHouse  Values============   : ");
		
		System.out.println("actstockLedgerRmaItemSecBalanceQtyAfterLoginToInvUser      :" + actstockLedgerRmaItemSecBalanceQtyAfterLoginToInvUser);
		System.out.println("expstockLedgerRmaItemSecBalanceQtyAfterLoginToInvUser      :" + expstockLedgerRmaItemSecBalanceQtyAfterLoginToInvUser);

		System.out.println("actstockLedgerRmaItemSecBalanceValueAfterLoginToInvUser    :" + actstockLedgerRmaItemSecBalanceValueAfterLoginToInvUser);
		System.out.println("expstockLedgerRmaItemSecBalanceValueAfterLoginToInvUser    :" + expstockLedgerRmaItemSecBalanceValueAfterLoginToInvUser);

		System.out.println("actstockLedgerrmaItemSecAvgRateAfterLoginToInvUser         :" + actstockLedgerrmaItemSecAvgRateAfterLoginToInvUser);
		System.out.println("expstockLedgerrmaItemSecAvgRateAfterLoginToInvUser         :" + expstockLedgerrmaItemSecAvgRateAfterLoginToInvUser);


		if (actstockLedgerRmaItemBalanceQty.equalsIgnoreCase(expstockLedgerRmaItemBalanceQty)
				&& actstockLedgerRmaItemBalanceValue.equalsIgnoreCase(expstockLedgerRmaItemBalanceValue)
				&& actstockLedgerrmaItemAvgRate.equalsIgnoreCase(expstockLedgerrmaItemAvgRate)
				&& actstockLedgerRmaItemHydBalanceQtyAfterLoginToInvUser.equalsIgnoreCase(expstockLedgerRmaItemHydBalanceQtyAfterLoginToInvUser)
				&& actstockLedgerRmaItemHydBalanceValueAfterLoginToInvUser.equalsIgnoreCase(expstockLedgerRmaItemHydBalanceValueAfterLoginToInvUser)
				&& actstockLedgerrmaItemHydAvgRateAfterLoginToInvUser.equalsIgnoreCase(expstockLedgerrmaItemHydAvgRateAfterLoginToInvUser)
				&& actstockLedgerRmaItemSecBalanceQtyAfterLoginToInvUser.equalsIgnoreCase(expstockLedgerRmaItemSecBalanceQtyAfterLoginToInvUser)
				&& actstockLedgerRmaItemSecBalanceValueAfterLoginToInvUser.equalsIgnoreCase(expstockLedgerRmaItemSecBalanceValueAfterLoginToInvUser)
				&& actstockLedgerrmaItemSecAvgRateAfterLoginToInvUser.equalsIgnoreCase(expstockLedgerrmaItemSecAvgRateAfterLoginToInvUser)) {

			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return true;
		} else {
			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	 @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td")
	  private static List<WebElement> Row1SalesInvoicevocherBodyList;
	 
	 @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td")
	  private static List<WebElement> Row2SalesInvoicevocherBodyList;
	 
	 @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td")
	  private static List<WebElement> Row3SalesInvoicevocherBodyList;
	 
	 @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td")
	  private static List<WebElement> Row4SalesInvoicevocherBodyList;
	 
	 @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td")
	  private static List<WebElement> Row8SalesInvoicevocherBodyList;
	 
	 @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td")
	  private static List<WebElement> Row9SalesInvoicevocherBodyList;
	 
	 @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td")
	  private static List<WebElement> Row5SalesInvoicevocherBodyList;
	
	public boolean checkSavingSalesInvoiceVoucherBrBatchItemWithHydWarehouse()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		
		
		
		click(financialsMenu);
		click(financialsTransactionMenu);
		click(financialsTransactionsSalesMenu);
		click(salesInvoiceVoucher);
		click(newBtn);
		checkValidationMessage("Screen opened");
		click(salesDatetxt);
		salesDatetxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		salesDatetxt.sendKeys(Keys.BACK_SPACE);
		enterText(salesDatetxt, "02/09/2021");
		tab(salesDatetxt);
		Thread.sleep(2000);
		enterText(salesAccounttxt, "Sales - Computers");
		tab(salesAccounttxt);
		Thread.sleep(2000);
		enterText(customerAccounttxt, "Bank");
		Thread.sleep(2000);
		click(customerAccounttxt);
		enterText(salesDepartmenttxt, "DEPT1");
		tab(salesDepartmenttxt);
		Thread.sleep(3000);		
		click(select1stRow_1stColumn);
		click(enterWarehousettxt);
		enterWarehousettxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		
		int warehousecount=svwareHouseListCount.size();
		System.err.println(warehousecount);
		ArrayList<String> warehouse = new ArrayList<String>();
		for(int i=0 ; i < warehousecount ;i++)
	  	{
	  		String data=svwareHouseListCount.get(i).getText();
	  		warehouse.add(data);
	  	}
		
		String actWarehouseList = warehouse.toString();
		String expWarehouseList = "[HYD, SEC]";
		
		System.out.println("actWarehouseList       :"+actWarehouseList);		
		System.out.println("expWarehouseList       :"+expWarehouseList);
		
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);		
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "BR-BATCH");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);
		
		String actAvailableStock      =enterAvailableStocktxt.getAttribute("value");
		String expAvailableStock	  ="12.000";
		
		
		System.out.println("actAvailableStock       :"+actAvailableStock);
		System.out.println("expAvailableStock       :"+expAvailableStock);
		
		
		tab(enterAvailableStocktxt);
		
		enterText(enterQuantitytxt, "5");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		Thread.sleep(2000);
		click(select1stRow_10thColumn);
		
		Thread.sleep(2000);
		
		String actRow1Baseqty		=batchPopUPRow1BaseQuantityTxt.getText();
		String expRow1Baseqty		="1.000";
		
		String actRow1Rate			=batchPopUPRow1RateTxt.getText();
		String expRow1Rate			="8.0101";
		
		String actRow2Baseqty		=batchPopUPRow2BaseQuantityTxt.getText();
		String expRow2Baseqty		="1.000";
		
		String actRow2Rate			=batchPopUPRow2RateTxt.getText();
		String expRow2Rate			="8.0404";
		
		String actRow3Baseqty		=batchPopUPRow3BaseQuantityTxt.getText();
		String expRow3Baseqty		="10.000";
		
		String actRow3Rate			=batchPopUPRow3RateTxt.getText();
		String expRow3Rate			="1000";
		
		System.out.println("actRow1Baseqty             :"+actRow1Baseqty);
		System.out.println("expRow1Baseqty             :"+expRow1Baseqty);
		
		System.out.println("actRow1Rate                :"+actRow1Rate);
		System.out.println("expRow1Rate                :"+expRow1Rate);
		
		System.out.println("actRow2Baseqty             :"+actRow2Baseqty);
		System.out.println("expRow2Baseqty             :"+expRow2Baseqty);
		
		System.out.println("actRow2Rate                :"+actRow2Rate);
		System.out.println("expRow2Rate                :"+expRow2Rate);
		
		System.out.println("actRow3Baseqty             :"+actRow3Baseqty);
		System.out.println("expRow3Baseqty             :"+expRow3Baseqty);
		
		System.out.println("actRow3Rate                :"+actRow3Rate);
		System.out.println("expRow3Rate                :"+expRow3Rate);
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		
		boolean actBatchpopupIsDisplaying = batchPickOnFIFOIcon.isDisplayed();
		boolean expBatchpopupIsDisplaying = true;
		
		System.out.println("Batch Popup isDisplaying  : "+actBatchpopupIsDisplaying+"  Value Expected  "+expBatchpopupIsDisplaying);
		click(batchPickOnFIFOIcon);
    	click(batchOkIcon);
    	
    	Thread.sleep(2000);
    	click(select4thRow_1stColumn);
    	Thread.sleep(2000);
    	
    	int row1List=Row1SalesInvoicevocherBodyList.size();
		System.err.println(row1List);
		ArrayList<String> row1SalesList = new ArrayList<String>();
		for(int i=0 ; i < row1List ;i++)
	  	{
	  		String data=Row1SalesInvoicevocherBodyList.get(i).getText();
	  		row1SalesList.add(data);
	  	}
		
		String actrow1SalesList = row1SalesList.toString();
		String exprow1SalesList = "[1, HYD, BR-BATCH, 1.000, 1.000, 100.0000, 100.0000, 0.00, 8.0101, 8.0101, 01082021PH, , ]";
		
		System.out.println("actrow1SalesList       :"+actrow1SalesList);		
		System.out.println("exprow1SalesList       :"+exprow1SalesList);
		
		int row2List=Row2SalesInvoicevocherBodyList.size();
		System.err.println(row2List);
		ArrayList<String> row2SalesList = new ArrayList<String>();
		for(int i=0 ; i < row2List ;i++)
	  	{
	  		String data=Row2SalesInvoicevocherBodyList.get(i).getText();
	  		row2SalesList.add(data);
	  	}
		
		String actrow2SalesList = row2SalesList.toString();
		String exprow2SalesList = "[2, HYD, BR-BATCH, 1.000, 1.000, 100.0000, 100.0000, 0.00, 8.0404, 8.0404, 04082021EXH, , ]";
		
		System.out.println("actrow2SalesList       :"+actrow2SalesList);		
		System.out.println("exprow2SalesList       :"+exprow2SalesList);
		
		
		int row3List=Row3SalesInvoicevocherBodyList.size();
		System.err.println(row3List);
		ArrayList<String> row3SalesList = new ArrayList<String>();
		for(int i=0 ; i < row3List ;i++)
	  	{
	  		String data=Row3SalesInvoicevocherBodyList.get(i).getText();
	  		row3SalesList.add(data);
	  	}
		
		String actrow3SalesList = row3SalesList.toString();
		String exprow3SalesList = "[3, HYD, BR-BATCH, 10.000, 3.000, 100.0000, 300.0000, 0.00, 1000.0000, 1000.0000, 01092021PH, , ]";
		
		System.out.println("actrow3SalesList       :"+actrow3SalesList);		
		System.out.println("exprow3SalesList       :"+exprow3SalesList);
		
    	Thread.sleep(2000);
    	
    	if(actWarehouseList.equalsIgnoreCase(expWarehouseList) && actAvailableStock.equalsIgnoreCase(expAvailableStock)
    			&& actRow1Baseqty.equalsIgnoreCase(expRow1Baseqty) && actRow1Rate.equalsIgnoreCase(expRow1Rate)
    			&& actRow2Baseqty.equalsIgnoreCase(expRow2Baseqty) && actRow2Rate.equalsIgnoreCase(expRow2Rate)
    			&& actRow3Baseqty.equalsIgnoreCase(expRow3Baseqty) && actRow3Rate.equalsIgnoreCase(expRow3Rate)
    			&& actrow1SalesList.equalsIgnoreCase(exprow1SalesList) && actrow2SalesList.equalsIgnoreCase(exprow2SalesList)
    			&& actrow3SalesList.equalsIgnoreCase(exprow3SalesList)) {
    		
    		return true;
    	}
    		else{
    			return false;
    		}
    		
    	}
	
	
			
	public boolean checkSavingSalesInvoiceVoucherBrBatchItemWithSECWarehouse() throws InterruptedException{	
		

		// 4th Row
		Thread.sleep(2000);

		//click(select4thRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(3000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "BR-BATCH");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);
		
		String actAvailableStock1      =enterAvailableStocktxt.getAttribute("value");
		String expAvailableStock1	   ="11.000";
		
		
		System.out.println("actAvailableStock1       :"+actAvailableStock1);
		System.out.println("expAvailableStock1       :"+expAvailableStock1);		
		
		
		tab(enterAvailableStocktxt);
		enterText(enterQuantitytxt, "5");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		click(select4thRow_10thColumn);
		
		Thread.sleep(2000);
		
		
		String actRow1Baseqty1		=batchPopUPRow1BaseQuantityTxt.getText();
		String expRow1Baseqty1		="1.000";
		
		String actRow1Rate1			=batchPopUPRow1RateTxt.getText();
		String expRow1Rate1			="8.0404";
		
		String actRow2Baseqty1		=batchPopUPRow2BaseQuantityTxt.getText();
		String expRow2Baseqty1		="10.000";
		
		String actRow2Rate1			=batchPopUPRow2RateTxt.getText();
		String expRow2Rate1			="2000";
		
	
		
		System.out.println("actRow1Baseqty1             :"+actRow1Baseqty1);
		System.out.println("expRow1Baseqty1             :"+expRow1Baseqty1);
		
		System.out.println("actRow1Rate1                :"+actRow1Rate1);
		System.out.println("expRow1Rate1                :"+expRow1Rate1);
		
		System.out.println("actRow2Baseqty1             :"+actRow2Baseqty1);
		System.out.println("expRow2Baseqty1             :"+expRow2Baseqty1);
		
		System.out.println("actRow2Rate1                :"+actRow2Rate1);
		System.out.println("expRow2Rate1                :"+expRow2Rate1);			
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		
		boolean actBatchpopupIsDisplaying1 = batchPickOnFIFOIcon.isDisplayed();
		boolean expBatchpopupIsDisplaying2 = true;
		
		System.out.println("Batch Popup isDisplaying  : "+actBatchpopupIsDisplaying1+"  Value Expected  "+expBatchpopupIsDisplaying2);
		click(batchPickOnFIFOIcon);
    	click(batchOkIcon);
    	
    	Thread.sleep(3000);
    	click(select6thRow_1stColumn);
    	Thread.sleep(3000);
    	
    	int row4List=Row4SalesInvoicevocherBodyList.size();
		System.err.println(row4List);
		ArrayList<String> row4SalesList = new ArrayList<String>();
		for(int i=0 ; i < row4List ;i++)
	  	{
	  		String data=Row4SalesInvoicevocherBodyList.get(i).getText();
	  		row4SalesList.add(data);
	  	}
		
		String actrow4SalesList = row4SalesList.toString();
		String exprow4SalesList = "[4, SEC, BR-BATCH, 1.000, 1.000, 100.0000, 100.0000, 0.00, 8.0404, 8.0404, 04082021EXH, , ]";
		
		System.out.println("actrow4SalesList       :"+actrow4SalesList);		
		System.out.println("exprow4SalesList       :"+exprow4SalesList);
		
		Thread.sleep(2000);
		//click(select6thRow_1stColumn);
		Thread.sleep(2000);
		
		int row5List=Row5SalesInvoicevocherBodyList.size();
		System.err.println(row5List);
		ArrayList<String> row5SalesList = new ArrayList<String>();
		for(int i=0 ; i < row5List ;i++)
	  	{
	  		String data=Row5SalesInvoicevocherBodyList.get(i).getText();
	  		row5SalesList.add(data);
	  	}
		
		String actrow5SalesList = row5SalesList.toString();
		String exprow5SalesList = "[5, SEC, BR-BATCH, 10.000, 4.000, 100.0000, 400.0000, 0.00, 2000.0000, 2000.0000, 01092021PH, , ]";
		
		System.out.println("actrow5SalesList       :"+actrow5SalesList);		
		System.out.println("exprow5SalesList       :"+exprow5SalesList);
		
		Thread.sleep(2000);
		
		if(actAvailableStock1.equalsIgnoreCase(expAvailableStock1) && actRow1Baseqty1.equalsIgnoreCase(expRow1Baseqty1) 
				&& actRow1Rate1.equalsIgnoreCase(expRow1Rate1) && actRow2Baseqty1.equalsIgnoreCase(expRow2Baseqty1) 
				&& actRow2Rate1.equalsIgnoreCase(expRow2Rate1) && actrow4SalesList.equalsIgnoreCase(exprow4SalesList)
				&& actrow5SalesList.equalsIgnoreCase(exprow5SalesList)){
			
			return true;
		}
			else{
				return false;
			}
			}
			
		
		
	public boolean checkSavingSalesInvoiceVoucherFifoItemWith2Warehouse() throws InterruptedException{ 	
		
		Thread.sleep(2000);

		//click(select6thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		
		
		Thread.sleep(2000);
		
		String actAvailableStockInHyd      =enterAvailableStocktxt.getAttribute("value");
		String expAvailableStockInHyd	   ="12.000";
		
		
		System.out.println("actAvailableStockInHyd        :"+actAvailableStockInHyd);
		System.out.println("expAvailableStockInHyd        :"+expAvailableStockInHyd);
		
		
		tab(enterAvailableStocktxt);
		enterText(enterQuantitytxt, "5");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		click(select6thRow_11thColumn);
		

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying + "  Value Expected  "
				+ expBinpopupIsDisplaying);

		click(binSearchBtn);
		Thread.sleep(3000);
		click(binAutoAllocateBtn);
		Thread.sleep(2000);
		click(binOkBtn);
		
		
		

		// 7th Row

		click(select7thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		
		
		Thread.sleep(2000);
		
		String actAvailableStockInSec      =enterAvailableStocktxt.getAttribute("value");
		String expAvailableStockInSec	   ="11.000";
		
		
		System.out.println("actAvailableStockInSec       :"+actAvailableStockInSec);
		System.out.println("expAvailableStockInSec       :"+expAvailableStockInSec);
		
		
		tab(enterAvailableStocktxt);
		enterText(enterQuantitytxt, "5");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		click(select7thRow_11thColumn);
		
	
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying1 = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying1 = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying1 + "  Value Expected  "
				+ expBinpopupIsDisplaying1);

		click(binSearchBtn);
		Thread.sleep(3000);
		click(binAutoAllocateBtn);
		Thread.sleep(2000);
		click(binOkBtn);
		
		Thread.sleep(2000);
		
		if(actAvailableStockInHyd.equalsIgnoreCase(expAvailableStockInHyd) && actBinpopupIsDisplaying==expBinpopupIsDisplaying
					&& actAvailableStockInSec.equalsIgnoreCase(expAvailableStockInSec) && actBinpopupIsDisplaying1==expBinpopupIsDisplaying1){
			
			return true;
		}
		else{
			return false;
			
		}
	}
		
	public boolean checkSavingSalesInvoiceVoucherRmaItemWithHYDWarehouse() throws InterruptedException{	
	

		// 8th Row
		Thread.sleep(2000);

		click(select8thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);
		tab(enterItemttxt);
		
		
		
		Thread.sleep(2000);
		
		String actAvailableStock4      =enterAvailableStocktxt.getAttribute("value");
		String expAvailableStock4	   ="12.000";
		
		
		System.out.println("actAvailableStock4       :"+actAvailableStock4);
		System.out.println("expAvailableStock4       :"+expAvailableStock4);	
		
		
		tab(enterAvailableStocktxt);
		enterText(enterQuantitytxt, "5");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		click(select8thRow_12thColumn);
		
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying + "  Value Expected  "
				+ expRMAPopupIsDisplaying);
		
		click(rmaSearchGridField);
		Thread.sleep(1000);
		
		int count = rmaSerialNos.size();
		
		ArrayList<String> SearchRMA = new ArrayList<String>();
		
		for (int i = 0; i < count; i++) 
		{
			String data = rmaSerialNos.get(i).getText();
			SearchRMA.add(data);
			
			if (data.equalsIgnoreCase("01092021PS") || data.equalsIgnoreCase("01082021PH1") || data.equalsIgnoreCase("01092021PS1") || data.equalsIgnoreCase("01092021PS2") || 
					data.equalsIgnoreCase("01092021PS3")) 
			{
				rmaSerialNoCheckBoxes.get(i).click();
			}
		}
		
		String actSearchRMASerialNoList = SearchRMA.toString();
		String expSearchRMASerialNoList = "[01092021PS, 01082021PH1, 01092021PS1, 01092021PS2, 01092021PS3, 01092021PS4, 01092021PS5, 01092021PS6, 01092021PS7, 01092021PS8, 01092021PS9, 25022021EXH1]"; 
		
		System.out.println("actSearchRMASerialNoList       :"+actSearchRMASerialNoList);
		System.out.println("expSearchRMASerialNoList       :"+expSearchRMASerialNoList);
		
		
		//click(searchRmaSelectAllGridCheckBox);
		click(searchRmaOkBtn);
		
		int serialcount	=rmaSerialNoList.size();
		ArrayList<String> rmaSerialnos = new ArrayList<String>();
		for (int i = 0; i < serialcount; i++) {
			String data=rmaSerialNoList.get(i).getText();
			rmaSerialnos.add(data);
		}
		
		String actRmaSerialNo 	=rmaSerialnos.toString();
		String expRmaSerialNo	="[01092021PS, 01082021PH1, 01092021PS1, 01092021PS2, 01092021PS3]";
		
		System.out.println("actRmaSerialNo       :"+actRmaSerialNo);
		System.out.println("expRmaSerialNo       :"+expRmaSerialNo);
		
	
		click(RMAPopupOkBtn);	
		Thread.sleep(2000);
		tab(enterRmatxt);
		
				
    	Thread.sleep(3000);
    	
    	
    	String actual8throwAvgRate 				= select8thRow_AvgRate.getText();
		String exp8throwAvgRate 				= "834.9155";
		String actual8throwOverallAvgRate		= select8thRow_OverAllAvgRate.getText();
		String exp8throwoverallAvgRate			= "1305.5580";
		
		String act8throwAvailableStock  		=select8thRow_AvailaleStock.getText();
		String exp8throwAvailableStock 			="12.000";
		
		System.out.println("actual8throwAvgRate             :"+actual8throwAvgRate);
		System.out.println("exp8throwAvgRate                :"+exp8throwAvgRate);
		
		System.out.println("actual8throwOverallAvgRate      :"+actual8throwOverallAvgRate);
		System.out.println("exp8throwoverallAvgRate         :"+exp8throwoverallAvgRate);
		
		System.out.println("act8throwAvailableStock         :"+act8throwAvailableStock);
		System.out.println("exp8throwAvailableStock         :"+exp8throwAvailableStock);
		
		if(actAvailableStock4.equalsIgnoreCase(expAvailableStock4) && actRMAPopupIsDisplaying==expRMAPopupIsDisplaying 
				&& actSearchRMASerialNoList.equalsIgnoreCase(expSearchRMASerialNoList) && actRmaSerialNo.equalsIgnoreCase(expRmaSerialNo)
				&& actual8throwAvgRate.equalsIgnoreCase(exp8throwAvgRate) && actual8throwOverallAvgRate.equalsIgnoreCase(exp8throwoverallAvgRate)
				&& act8throwAvailableStock.equalsIgnoreCase(exp8throwAvailableStock)){
			return true;
			
		}else{
			return false;
		}
	}
	
	public boolean checkSavingSalesInvoiceVoucherRmaItemWithSECWarehouse() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{

		// 9th Row
		Thread.sleep(1000);
		//getAction().moveToElement(select9thRow_1stColumn).click().perform();
		click(select9thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);
		tab(enterItemttxt);
		
		
		Thread.sleep(2000);
		
		String actAvailableStock5      =enterAvailableStocktxt.getAttribute("value");
		String expAvailableStock5	   ="11.000";
		
		
		System.out.println("actAvailableStock5       :"+actAvailableStock5);
		System.out.println("expAvailableStock5       :"+actAvailableStock5);	
		
		
		
		tab(enterAvailableStocktxt);
		enterText(enterQuantitytxt, "5");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		click(select9thRow_12thColumn);
		
		

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying1 = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying1 = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying1 + "  Value Expected  "
				+ expRMAPopupIsDisplaying1);
		
		click(rmaSearchGridField);
		Thread.sleep(1000);
		
		int count = rmaSerialNos.size();
		
		ArrayList<String> SearchRMA = new ArrayList<String>();
		
		for (int i = 0; i < count; i++) 
		{
			String data = rmaSerialNos.get(i).getText();
			SearchRMA.add(data);
			
			if (data.equalsIgnoreCase("01092021PH9") || data.equalsIgnoreCase("01092021PH10") || data.equalsIgnoreCase("01092021PH11") || data.equalsIgnoreCase("01092021PH12") || 
					data.equalsIgnoreCase("01092021PH13")) 
			{
				rmaSerialNoCheckBoxes.get(i).click();
			}
		}
		
		String actSearchRMASerialNoList = SearchRMA.toString();
		String expSearchRMASerialNoList = "[01092021PH9, 01092021PH10, 01092021PH11, 01092021PH12, 01092021PH13, 01092021PH14, 01092021PH15, 01092021PH16, 01092021PH17, 01092021PH18, 04082021EXH1]"; 
		
		System.out.println("actSearchRMASerialNoList       :"+actSearchRMASerialNoList);
		System.out.println("expSearchRMASerialNoList       :"+expSearchRMASerialNoList);
		
		
		//click(searchRmaSelectAllGridCheckBox);
		click(searchRmaOkBtn);
		
		int serialcount	=rmaSerialNoList.size();
		ArrayList<String> rmaSerialnos = new ArrayList<String>();
		for (int i = 0; i < serialcount; i++) {
			String data=rmaSerialNoList.get(i).getText();
			rmaSerialnos.add(data);
		}
		
		String actRmaSerialNo 	=rmaSerialnos.toString();
		String expRmaSerialNo	="[01092021PH9, 01092021PH10, 01092021PH11, 01092021PH12, 01092021PH13]";
		
		System.out.println("actRmaSerialNo       :"+actRmaSerialNo);
		System.out.println("expRmaSerialNo       :"+expRmaSerialNo);		
		
		
		click(RMAPopupOkBtn);	
		Thread.sleep(1000);
		tab(enterRmatxt);
		
		
		Thread.sleep(3000);
    	
    	
    	String actual9throwAvgRate 				= select9thRow_AvgRate.getText();
		String exp9throwAvgRate 				= "1819.0042";
		String actual9throwOverallAvgRate		= select9thRow_OverAllAvgRate.getText();
		String exp9throwoverallAvgRate			= "1305.5580";
		
		String act9throwAvailableStock  		=select9thRow_AvailaleStock.getText();
		String exp9throwAvailableStock 			="11.000";
		
		System.out.println("actual9throwAvgRate             :"+actual9throwAvgRate);
		System.out.println("exp9throwAvgRate                :"+exp9throwAvgRate);
		
		System.out.println("actual9throwOverallAvgRate      :"+actual9throwOverallAvgRate);
		System.out.println("exp9throwoverallAvgRate         :"+exp9throwoverallAvgRate);
		
		System.out.println("act9throwAvailableStock         :"+act9throwAvailableStock);
		System.out.println("exp9throwAvailableStock         :"+exp9throwAvailableStock);
		
				
	
		Thread.sleep(2000);
		click(voucherSaveBtn);

		HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) {
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : 35");
		expMessage.add("Saving in background.");

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage);

		if(actMessage.equals(expMessage) && actAvailableStock5.equalsIgnoreCase(expAvailableStock5) && actRMAPopupIsDisplaying1==expRMAPopupIsDisplaying1 
				&& actSearchRMASerialNoList.equalsIgnoreCase(expSearchRMASerialNoList) && actRmaSerialNo.equalsIgnoreCase(expRmaSerialNo)
				&& actual9throwAvgRate.equalsIgnoreCase(exp9throwAvgRate) && actual9throwOverallAvgRate.equalsIgnoreCase(exp9throwoverallAvgRate)
				&& act9throwAvailableStock.equalsIgnoreCase(exp9throwAvailableStock)){	
			click(new_CloseBtn);			
			Thread.sleep(2000);
			click(voucherhomeCloseBtn);
			return true;
			
		} else {
			click(new_CloseBtn);			
			Thread.sleep(2000);
			click(voucherhomeCloseBtn);
			return false;

		}

	}
	
	public boolean checkStockLedgerReportForBrBatchItemAfterSavingSalesInvoiceVocher()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		
		
		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryStockLedgerMenu);
		Thread.sleep(3000);

		int rowcount = stockLedgerHometableRowCount.size();

		System.out.println(rowcount);

		for (int i = 0; i < rowcount; i++) {
			String actName = ledgerHometableItemNamesList.get(i).getText();

			System.out.println(actName);

			if (actName.equalsIgnoreCase("BR-BATCH")) {
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);

		String actStockLedgerBrBatchIssuedrateTxt = getText(stockLedgerBrBatchAfterSavingSalesInvoiceIssuedrateTxt);
		String expStockLedgerBrBatchIssuedrateTxt = "2,000.0000";
		
		
		
		
		System.out.println("actStockLedgerBrBatchIssuedrateTxt      :" + actStockLedgerBrBatchIssuedrateTxt);
		System.out.println("expStockLedgerBrBatchIssuedrateTxt      :" + expStockLedgerBrBatchIssuedrateTxt);
		
		
	

		if (actStockLedgerBrBatchIssuedrateTxt.equalsIgnoreCase(expStockLedgerBrBatchIssuedrateTxt)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}
	
	public boolean checkFinanacialLedgerReportFifoItemAfterSavingSalesInvoiceVocher() // NOt Implemented
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == false) {
			click(fifoItemCheckBox);
		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == true) {
			click(rmaItemCheckBox);

		}

		
		dropDown(sl_DateOptionDropdown, "1");	
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);

		String actStockLedgerFifoItemIssuedrateTxt = getText(stockLedgerFifoItemAfterSavingSalesInvoiceIssuedrateTxt);
		String expStockLedgerFifoItemIssuedrateTxt = "1,000.0000";
		
		
		String actstockLedgerFifoItemBalanceQty    = getText(stockLedgerFifoItemBalanceQtyAfterSavingPurchase);
		String expstockLedgerFifoItemBalanceQty    = "13.0000";

		String actstockLedgerFifoItemBalanceValue  = getText(stockLedgerFifoItemBalanceValueAfterSavingPurchase);
		String expstockLedgerFifoItemBalanceValue  = "30,041.0807";

		String actstockLedgerFifoItemAvgRate       = getText(stockLedgerFifoItemAvgRateAfterSavingPurchase);
		String expstockLedgerFifoItemAvgRate       = "1,201.6432";
		
		System.out.println("actStockLedgerFifoItemIssuedrateTxt      :" + actStockLedgerFifoItemIssuedrateTxt);
		System.out.println("expStockLedgerFifoItemIssuedrateTxt      :" + expStockLedgerFifoItemIssuedrateTxt);		

		System.out.println("actstockLedgerFifoItemBalanceQty         :" + actstockLedgerFifoItemBalanceQty);
		System.out.println("expstockLedgerFifoItemBalanceQty         :" + expstockLedgerFifoItemBalanceQty);

		System.out.println("actstockLedgerFifoItemBalanceValue       :" + actstockLedgerFifoItemBalanceValue);
		System.out.println("expstockLedgerFifoItemBalanceValue       :" + expstockLedgerFifoItemBalanceValue);

		System.out.println("actstockLedgerFifoItemAvgRate            :" + actstockLedgerFifoItemAvgRate);
		System.out.println("expstockLedgerFifoItemAvgRate            :" + expstockLedgerFifoItemAvgRate);

		if (actstockLedgerFifoItemBalanceQty.equalsIgnoreCase(expstockLedgerFifoItemBalanceQty)
				&& actstockLedgerFifoItemBalanceValue.equalsIgnoreCase(expstockLedgerFifoItemBalanceValue)
				&& actstockLedgerFifoItemAvgRate.equalsIgnoreCase(expstockLedgerFifoItemAvgRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}

	public boolean checkStockLedgerReportForRmaItemAfterSavingSalesInvoiceVocher()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == true) {
			click(fifoItemCheckBox);

		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == false) {
			click(rmaItemCheckBox);

		}

	
		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);
		
		
		String actStockLedgerRmaItemIssuedrateTxt = getText(stockLedgerRmaItemAfterSavingSalesInvoiceIssuedrateTxt);
		String expStockLedgerRmaItemIssuedrateTxt = "1,819.0042";

		
		System.out.println("actStockLedgerRmaItemIssuedrateTxt      :" + actStockLedgerRmaItemIssuedrateTxt);
		System.out.println("expStockLedgerRmaItemIssuedrateTxt      :" + expStockLedgerRmaItemIssuedrateTxt);
		

	
		if (actStockLedgerRmaItemIssuedrateTxt.equalsIgnoreCase(expStockLedgerRmaItemIssuedrateTxt)) {

			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return true;
		} else {
			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	@FindBy(xpath="//input[@id='id_header_2']")
	private static WebElement  stockTransferDatetxt;
	
	@FindBy(xpath="//select[@id='id_header_31']")
	private static WebElement  stockTransferIssuestxt;
	
	@FindBy(xpath="//input[@id='id_header_86']")
	private static WebElement  stockTransferWarehousetxt;
	
	@FindBy(xpath="//input[@id='id_header_86']")
	private static WebElement  Warehouse1txt;
	
	@FindBy(xpath="//input[@id='id_header_87']")
	private static WebElement  Warehouse2txt;
	
	@FindBy(xpath="//input[@id='id_body_87']")
	private static WebElement  enterWarehousettxt2;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[7]")
	private static WebElement  select1stRow_AvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[7]")
	private static WebElement  select2ndRow_AvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[8]")
	private static WebElement  select1stRow_OverAllAvgRate;
	
	
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[8]")
	private static WebElement  select2ndRow_OverAllAvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[9]")
	private static WebElement  select1stRow_BatchText;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[12]")
	private static WebElement  select2ndRow_11thColumn;
	
	
	public boolean checkLoginToSUUserAndSavingStockTransferVocher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(3000);
		click(userNameDisplay);
		Thread.sleep(1000);
		click(logoutOption);		
		Thread.sleep(3000);
		
		String unamelt="su";

		String pawslt="su";
		
		LoginPage.enterUserName(unamelt);	     
        LoginPage.enterPassword(pawslt);
        
		
		 Select company = new Select(companyDropDownList);
		 company.selectByValue("108");
		 Thread.sleep(2000);
		 
  
        LoginPage.clickOnSignInBtn();  
        
        Thread.sleep(3000);
        
        System.out.println("Signing");
        
        // Login User Name is Verify
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		String actUserInfo                      = userNameDisplay.getText();
		String expUserInfo                      = "SU";	

		System.out.println("Login With Valid User and Check User Name Display Value Actual : ." + actUserInfo + ". Value Expected : " + expUserInfo);
		
		
		// Login Company Name is Verify
		Thread.sleep(1000);
		click(companyLogo);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
		String actGetLoginCompanyNameInformation = companyName.getText();
		String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
		String expGetLoginCompanyName            = "Maintain Quantity a";   
		companyLogo.click();
		
		System.out.println("Company Name Display Value Actual                : ." + actGetLoginCompanyName +  ". Value Expected : " + expGetLoginCompanyName);
		


	Thread.sleep(2000);
	click(inventoryMenu);
	click(inventoryTransactionsMenu);
	click(inventoryTransactionsStockMenu);
	click(inventoryTransactionsStockMenuStockTransfersMenu);	
	click(newBtn);
	
	checkValidationMessage("Screen opened");
	
	click(stockTransferDatetxt);
	stockTransferDatetxt.sendKeys(Keys.CONTROL, "a");
	Thread.sleep(2000);
	stockTransferDatetxt.sendKeys(Keys.BACK_SPACE);
	enterText(stockTransferDatetxt, "02/09/2021");
	tab(stockTransferDatetxt);
	tab(stockTransferIssuestxt);
	Thread.sleep(1000);
	enterText(stockTransferWarehousetxt, "HYD");
	Thread.sleep(1000);
	tab(stockTransferWarehousetxt);
	Thread.sleep(2000);
	
	
	click(select1stRow_1stColumn);
	enterText(enterWarehousettxt2, "Test");
	Thread.sleep(2000);
	
	

	tab(enterWarehousettxt2);
	enterText(enterItemttxt, "BR-BATCH");
	Thread.sleep(2000);
	tab(enterItemttxt);
	
	

	
	enterText(enterQuantitytxt, "3");
	tab(enterQuantitytxt);
	enterText(enterRatetxt, "100");
	tab(enterRatetxt);
	Thread.sleep(2000);
	
	
	String actRow1AvgRate				=select1stRow_AvgRate.getText();
	String expRow1AvgRate				="0.0000";
	
	String actRow1OverAllAvgRate		=select1stRow_OverAllAvgRate.getText();
	String expRow1OverAllAvgRate		="0.0000";
	
	
	
	System.out.println("actRow1AvgRate             :"+actRow1AvgRate);
	System.out.println("expRow1AvgRate             :"+expRow1AvgRate);
	
	System.out.println("actRow1OverAllAvgRate      :"+actRow1OverAllAvgRate);
	System.out.println("expRow1OverAllAvgRate      :"+expRow1OverAllAvgRate);
	
	
	click(select1stRow_8thColumn);
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
	
	boolean actBatchpopupIsDisplaying = batchPickOnFIFOIcon.isDisplayed();
	boolean expBatchpopupIsDisplaying = true;
	
	System.out.println("Batch Popup isDisplaying  : "+actBatchpopupIsDisplaying+"  Value Expected  "+expBatchpopupIsDisplaying);
	click(batchPickOnFIFOIcon);
	click(batchOkIcon);
	
	Thread.sleep(2000);
	
	
	String actRow1AvgRate1				=select1stRow_AvgRate.getText();
	String expRow1AvgRate1				="1000.0000";
	
	String actRow1OverAllAvgRate1		=select1stRow_OverAllAvgRate.getText();
	String expRow1OverAllAvgRate1		="1000.0000";
	
	String actRow1BatchText				=select1stRow_BatchText.getText();
	String expRow1BatchText				="01092021PH";
	
	
	
	System.out.println("actRow1AvgRate1             :"+actRow1AvgRate1);
	System.out.println("expRow1AvgRate1             :"+expRow1AvgRate1);
	
	System.out.println("actRow1OverAllAvgRate1      :"+actRow1OverAllAvgRate1);
	System.out.println("expRow1OverAllAvgRate1      :"+expRow1OverAllAvgRate1);
	
	System.out.println("actRow1BatchText            :"+actRow1BatchText);
	System.out.println("expRow1BatchText            :"+expRow1BatchText);



	click(select2ndRow_1stColumn);
	
	enterWarehousettxt2.sendKeys(Keys.CONTROL, "a");
	Thread.sleep(2000);
	enterWarehousettxt2.sendKeys(Keys.BACK_SPACE);	
	enterText(enterWarehousettxt2, "Test");
	Thread.sleep(2000);
	tab(enterWarehousettxt2);
	enterText(enterItemttxt, "WA-RMA");
	Thread.sleep(2000);
	tab(enterItemttxt);
	
	

	
	enterText(enterQuantitytxt, "3");
	tab(enterQuantitytxt);
	enterText(enterRatetxt, "100");
	tab(enterRatetxt);
	Thread.sleep(2000);	
	
	click(select2ndRow_11thColumn);
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

	boolean actRMAPopupIsDisplaying1 = rmaScreenTitle.isDisplayed();
	boolean expRMAPopupIsDisplaying1 = true;

	System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying1 + "  Value Expected  "
			+ expRMAPopupIsDisplaying1);
	
	click(rmaSearchGridField);
	Thread.sleep(1000);
	
	int count = rmaSerialNos.size();
	
	ArrayList<String> SearchRMA = new ArrayList<String>();
	
	for (int i = 0; i < count; i++) 
	{
		String data = rmaSerialNos.get(i).getText();
		SearchRMA.add(data);
		
		if (data.equalsIgnoreCase("01092021PS4") || data.equalsIgnoreCase("01092021PS5") || data.equalsIgnoreCase("01092021PS6")) 
		{
			rmaSerialNoCheckBoxes.get(i).click();
		}
	}
	
	String actSearchRMASerialNoList = SearchRMA.toString();
	String expSearchRMASerialNoList = "[01092021PS4, 01092021PS5, 01092021PS6, 01092021PS7, 01092021PS8, 01092021PS9, 25022021EXH1]"; 
	
	System.out.println("actSearchRMASerialNoList       :"+actSearchRMASerialNoList);
	System.out.println("expSearchRMASerialNoList       :"+expSearchRMASerialNoList);
	
	
	//click(searchRmaSelectAllGridCheckBox);
	click(searchRmaOkBtn);
	
	int serialcount	=rmaSerialNoList.size();
	ArrayList<String> rmaSerialnos = new ArrayList<String>();
	for (int i = 0; i < serialcount; i++) {
		String data=rmaSerialNoList.get(i).getText();
		rmaSerialnos.add(data);
	}
	
	String actRmaSerialNo 	=rmaSerialnos.toString();
	String expRmaSerialNo	="[01092021PS4, 01092021PS5, 01092021PS6]";
	
	System.out.println("actRmaSerialNo       :"+actRmaSerialNo);
	System.out.println("expRmaSerialNo       :"+expRmaSerialNo);
	
	
	click(RMAPopupOkBtn);	
	Thread.sleep(1000);
	tab(enterRmatxt);
	
	
	Thread.sleep(3000);
	
	
	String actRow2AvgRate				=select2ndRow_AvgRate.getText();
	String expRow2AvgRate				="834.9155";
	
	String actRow2OverAllAvgRate		=select2ndRow_OverAllAvgRate.getText();
	String expRow2OverAllAvgRate		="1201.9204";
	
	
	
	System.out.println("actRow2AvgRate             		  :"+actRow2AvgRate);
	System.out.println("expRow2AvgRate             		  :"+expRow2AvgRate);
	
	System.out.println("actRow2OverAllAvgRate             :"+actRow2OverAllAvgRate);
	System.out.println("expRow2OverAllAvgRate             :"+expRow2OverAllAvgRate);
	
	// 3rd Row
	
	click(select3rdRow_1stColumn);
	
	enterWarehousettxt2.sendKeys(Keys.CONTROL, "a");
	Thread.sleep(2000);
	enterWarehousettxt2.sendKeys(Keys.BACK_SPACE);	
	enterWarehousettxt2.sendKeys(Keys.SPACE);
	Thread.sleep(1000);
	enterText(enterWarehousettxt2, "TEST");
	Thread.sleep(2000);
	tab(enterWarehousettxt2);
	enterText(enterItemttxt, "FIFO-BIN");
	Thread.sleep(2000);
	tab(enterItemttxt);
	
	

	
	enterText(enterQuantitytxt, "3");
	tab(enterQuantitytxt);
	enterText(enterRatetxt, "100");
	tab(enterRatetxt);
	Thread.sleep(2000);	
	
	click(select3rdRow_9thColumn);
	Thread.sleep(2000);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

	boolean actBinpopupIsDisplaying = binSearchTxt.isDisplayed();
	boolean expBinpopupIsDisplaying = true;

	System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying + "  Value Expected  "
			+ expBinpopupIsDisplaying);

	click(binSearchBtn);
	Thread.sleep(3000);
	click(binPopup1stRow);
	Thread.sleep(1500);
	click(binPickBtn);
	/*Thread.sleep(2000);
	click(binAutoAllocateBtn);*/
	Thread.sleep(2000);
	click(binOkBtn);
	
	Thread.sleep(2000);
	
	tab(bin1Txt);
	
	//bin2Txt.sendKeys(Keys.SPACE);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

	boolean actBinpopupIsDisplaying1 = binSearchTxt.isDisplayed();
	boolean expBinpopupIsDisplaying1 = true;

	System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying1 + "  Value Expected  "
			+ expBinpopupIsDisplaying1);

	click(binSearchBtn);
	Thread.sleep(3000);
	click(binPopup1stRow);
	Thread.sleep(1500);
	click(binPickBtn);
	//click(binAutoAllocateBtn);
	Thread.sleep(2000);
	click(binOkBtn);
	
	//bin2Txt.sendKeys(Keys.SPACE);
	
	
	Thread.sleep(2000);
	click(voucherSaveBtn);

	HashSet<String> actMessage = new HashSet<String>();

	for (int i = 0; i < 2; i++) {
		String data = checkValidationMessage("");
		actMessage.add(data);
	}

	HashSet<String> expMessage = new HashSet<String>();

	expMessage.add("Voucher saved successfully : 70");
	expMessage.add("Saving in background.");

	System.out.println("Actual Message    : " + actMessage);
	System.out.println("Expected Message  : " + expMessage);

	if(actMessage.equals(expMessage) && actRow1AvgRate.equalsIgnoreCase(expRow1AvgRate) && actRMAPopupIsDisplaying1==expRMAPopupIsDisplaying1 
			&& actRow1OverAllAvgRate.equalsIgnoreCase(expRow1OverAllAvgRate) && actRow1AvgRate1.equalsIgnoreCase(expRow1AvgRate1)
			&& actRow1OverAllAvgRate1.equalsIgnoreCase(expRow1OverAllAvgRate1) && actRow1BatchText.equalsIgnoreCase(expRow1BatchText)
			&& actSearchRMASerialNoList.equalsIgnoreCase(expSearchRMASerialNoList) && actRmaSerialNo.equalsIgnoreCase(expRmaSerialNo)
			&& actRow2AvgRate.equalsIgnoreCase(expRow2AvgRate) && actRow2OverAllAvgRate.equalsIgnoreCase(expRow2OverAllAvgRate)){	
		click(new_CloseBtn);			
		Thread.sleep(2000);
		click(voucherhomeCloseBtn);
		return true;
		
	} else {
		click(new_CloseBtn);			
		Thread.sleep(2000);
		click(voucherhomeCloseBtn);
		return false;

	  }
	}	
	
	
	public boolean validteItemQueryStocksInBatchItemAfterSavingStockTransferVoucher() throws InterruptedException {

		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(itemqueryMenu);
		click(itemTxt);
		Thread.sleep(2000);
		enterText(itemTxt, "BR-BATCH");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "15.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "1,267.7481";

		String actlstockValueTxt 	 = getText(stockValueTxt);
		String expstockValueTxt		 = "19,016.2222";

		System.out.println("=====Stocks In Batch Items======");

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock           :" + actlCurrentStock);
		System.out.println("expCurrentStock            :" + expCurrentStock);

		System.out.println("actlavgStockRate           :" + actlavgStockRate);
		System.out.println("expavgStockRate            :" + expavgStockRate);

		System.out.println("actlstockValueTxt          :" + actlstockValueTxt);
		System.out.println("expstockValueTxt           :" + expstockValueTxt);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)) {
			return true;
		} else {
			return false;
		}

	}

	public boolean validteItemQueryStocksInFifoItemAfterSavingStockTransferVoucher() throws InterruptedException {

		click(itemTxt);
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.BACK_SPACE);
		enterText(itemTxt, "FIFO-BIN");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "15.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "1,666.6667";

		String actlstockValueTxt     = getText(stockValueTxt);
		String expstockValueTxt      = "25,000.0000";

		System.out.println("=====Stocks In FIFO Items======");

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock           :" + actlCurrentStock);
		System.out.println("expCurrentStock            :" + expCurrentStock);

		System.out.println("actlavgStockRate           :" + actlavgStockRate);
		System.out.println("expavgStockRate            :" + expavgStockRate);

		System.out.println("actlstockValueTxt          :" + actlstockValueTxt);
		System.out.println("expstockValueTxt           :" + expstockValueTxt);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)) {
			return true;
		} else {
			return false;
		}

	}

	public boolean validteItemQueryStocksInWAItemAfterSavingStockTransferVoucher() throws InterruptedException {

		click(itemTxt);
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.BACK_SPACE);
		enterText(itemTxt, "WA-RMA");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "15.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "1,128.5194";

		String actlstockValueTxt     = getText(stockValueTxt);
		String expstockValueTxt      = "16,927.7910";

		ScrollToElement(balanceByWarehouseBtn);
		click(balanceByWarehouseBtn);

		Thread.sleep(3000);

		int count = row1BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow1 = new ArrayList<String>();

		for (int i = 0; i < count; i++) {
			String data = row1BalanceByWareHouseList.get(i).getText();
			warehouseRow1.add(data);
		}

		String actwarehouseRow1List = warehouseRow1.toString();
		String expWarehouseRow1List = "[1, HYD, 4.000, 0.000, 4.000, 834.9155, 3,339.6618]";

		int count1 = row2BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow2 = new ArrayList<String>();

		for (int i = 0; i < count1; i++) {
			String data = row2BalanceByWareHouseList.get(i).getText();
			warehouseRow2.add(data);
		}

		String actwarehouseRow2List = warehouseRow2.toString();
		String expWarehouseRow2List = "[2, SEC, 6.000, 0.000, 6.000, 1,819.0042, 10,914.0250]";

		int count2 = row3BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow3 = new ArrayList<String>();

		for (int i = 0; i < count2; i++) {
			String data = row3BalanceByWareHouseList.get(i).getText();
			warehouseRow3.add(data);
		}

		String actwarehouseRow3List = warehouseRow3.toString();
		String expWarehouseRow3List = "[3, TEST, 5.000, 0.000, 5.000, 505.4718, 2,527.3588]";
		
		

		System.out.println("=====Stocks In WA Items======");

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock           :" + actlCurrentStock);
		System.out.println("expCurrentStock            :" + expCurrentStock);

		System.out.println("actlavgStockRate           :" + actlavgStockRate);
		System.out.println("expavgStockRate            :" + expavgStockRate);

		System.out.println("actlstockValueTxt          :" + actlstockValueTxt);
		System.out.println("expstockValueTxt           :" + expstockValueTxt);

		System.out.println("===========Balance By Warehouse=======");

		System.out.println(actwarehouseRow1List);
		System.out.println(expWarehouseRow1List);

		System.out.println(actwarehouseRow2List);
		System.out.println(expWarehouseRow2List);

		System.out.println(actwarehouseRow3List);
		System.out.println(expWarehouseRow3List);

		
	
		
		
		
		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)
				&& actwarehouseRow1List.equalsIgnoreCase(expWarehouseRow1List)
				&& actwarehouseRow2List.equalsIgnoreCase(expWarehouseRow2List)
				&& actwarehouseRow3List.equalsIgnoreCase(expWarehouseRow3List))
		{
			return true;
		} else {
			return false;
		}

	}
	
	
	
	
	@FindBy(xpath = "(//tr[@id='trRender_12'])[1]/td[6]")
	private static WebElement stockLedgerBrBatchReceivedRateAfterSavingStockTransfer;
	
	@FindBy(xpath = "(//tr[@id='trRender_11'])[1]/td[8]")
	private static WebElement stockLedgerBrBatchIssuedRateAfterSavingStockTransfer;	
	
	@FindBy(xpath = "(//tr[@id='trRender_12'])[1]/td[13]")
	private static WebElement stockLedgerBrBatchAvgRateAfterSavingStockTransfer;
	
	@FindBy(xpath = "(//tr[@id='trRender_12'])[1]/td[9]")
	private static WebElement stockLedgerBrBatchBalanceQtyAfterSavingStockTransfer;
	
	@FindBy(xpath = "(//tr[@id='trRender_12'])[1]/td[12]")
	private static WebElement stockLedgerBrBatchBalanceValueAfterSavingStockTransfer;


	
	
	
	public boolean checkStockLedgerReportForBrBatchItemAfterStockTransferVoucher()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryStockLedgerMenu);
		Thread.sleep(3000);

		int rowcount = stockLedgerHometableRowCount.size();

		System.out.println(rowcount);

		for (int i = 0; i < rowcount; i++) {
			String actName = ledgerHometableItemNamesList.get(i).getText();

			System.out.println(actName);

			if (actName.equalsIgnoreCase("BR-BATCH")) {
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);

		String actStockLedgerBrBatchReceivedRate 	= getText(stockLedgerBrBatchReceivedRateAfterSavingStockTransfer);
		String expStockLedgerBrBatchReceivedRate	= "1,000.0000";
		
		String actStockLedgerBrBatchIssuedRate 		= getText(stockLedgerBrBatchIssuedRateAfterSavingStockTransfer);
		String expStockLedgerBrBatchIssuedRate		= "1,000.0000";
		
		
		String actStockLedgerBrBatchBalanceQty 		= getText(stockLedgerBrBatchBalanceQtyAfterSavingStockTransfer);
		String expStockLedgerBrBatchBalanceQty 	    = "15.0000";

		String actStockLedgerBrBatchBalanceValue    = getText(stockLedgerBrBatchBalanceValueAfterSavingStockTransfer);
		String expStockLedgerBrBatchBalanceValue    = "19,016.2222";

		String actStockLedgerBrBatchAvgRate 		= getText(stockLedgerBrBatchAvgRateAfterSavingStockTransfer);
		String expStockLedgerBrBatchAvgRate			= "1,267.7481";

		
		
		System.out.println("actStockLedgerBrBatchReceivedRate    :" + actStockLedgerBrBatchReceivedRate);
		System.out.println("expStockLedgerBrBatchReceivedRate    :" + expStockLedgerBrBatchReceivedRate);

		System.out.println("actStockLedgerBrBatchIssuedRate      :" + actStockLedgerBrBatchIssuedRate);
		System.out.println("expStockLedgerBrBatchIssuedRate      :" + expStockLedgerBrBatchIssuedRate);
		
		System.out.println("actStockLedgerBrBatchBalanceQty      :" + actStockLedgerBrBatchBalanceQty);
		System.out.println("expStockLedgerBrBatchBalanceQty      :" + expStockLedgerBrBatchBalanceQty);


		System.out.println("actStockLedgerBrBatchBalanceValue    :" + actStockLedgerBrBatchBalanceValue);
		System.out.println("expStockLedgerBrBatchBalanceValue    :" + expStockLedgerBrBatchBalanceValue);

		System.out.println("actStockLedgerBrBatchAvgRate         :" + actStockLedgerBrBatchAvgRate);
		System.out.println("expStockLedgerBrBatchAvgRate         :" + expStockLedgerBrBatchAvgRate);

		if (actStockLedgerBrBatchBalanceQty.equalsIgnoreCase(expStockLedgerBrBatchBalanceQty)
				&& actStockLedgerBrBatchBalanceValue.equalsIgnoreCase(expStockLedgerBrBatchBalanceValue)
				&& actStockLedgerBrBatchAvgRate.equalsIgnoreCase(expStockLedgerBrBatchAvgRate) && actStockLedgerBrBatchReceivedRate.equalsIgnoreCase(expStockLedgerBrBatchReceivedRate)
				&& actStockLedgerBrBatchIssuedRate.equalsIgnoreCase(expStockLedgerBrBatchIssuedRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}
	
	
	@FindBy(xpath = "(//tr[@id='trRender_13'])[1]/td[13]")
	private static WebElement stockLedgerFifoItemAvgRateAfterSavingStockTransfer;
	
	@FindBy(xpath = "(//tr[@id='trRender_13'])[1]/td[9]")
	private static WebElement stockLedgerFifoItemBalanceQtyAfterSavingStockTransfer;
	
	@FindBy(xpath = "(//tr[@id='trRender_13'])[1]/td[12]")
	private static WebElement stockLedgerFifoItemBalanceValueAfterSavingStockTransfer;

	public boolean checkStockLedgerReportForFifoItemAfterStockTransferVoucher()  
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == false) {
			click(fifoItemCheckBox);
		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == true) {
			click(rmaItemCheckBox);

		}

		
		dropDown(sl_DateOptionDropdown, "1");	
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);

		String actstockLedgerFifoItemBalanceQty   = getText(stockLedgerFifoItemBalanceQtyAfterSavingStockTransfer);
		String expstockLedgerFifoItemBalanceQty   = "15.0000";

		String actstockLedgerFifoItemBalanceValue = getText(stockLedgerFifoItemBalanceValueAfterSavingStockTransfer);
		String expstockLedgerFifoItemBalanceValue = "25,000.0000";

		String actstockLedgerFifoItemAvgRate      = getText(stockLedgerFifoItemAvgRateAfterSavingStockTransfer);
		String expstockLedgerFifoItemAvgRate      = "1,666.6667";

		System.out.println("actstockLedgerFifoItemBalanceQty        :" + actstockLedgerFifoItemBalanceQty);
		System.out.println("expstockLedgerFifoItemBalanceQty        :" + expstockLedgerFifoItemBalanceQty);

		System.out.println("actstockLedgerFifoItemBalanceValue      :" + actstockLedgerFifoItemBalanceValue);
		System.out.println("expstockLedgerFifoItemBalanceValue      :" + expstockLedgerFifoItemBalanceValue);

		System.out.println("actstockLedgerFifoItemAvgRate           :" + actstockLedgerFifoItemAvgRate);
		System.out.println("expstockLedgerFifoItemAvgRate           :" + expstockLedgerFifoItemAvgRate);

		if (actstockLedgerFifoItemBalanceQty.equalsIgnoreCase(expstockLedgerFifoItemBalanceQty)
				&& actstockLedgerFifoItemBalanceValue.equalsIgnoreCase(expstockLedgerFifoItemBalanceValue)
				&& actstockLedgerFifoItemAvgRate.equalsIgnoreCase(expstockLedgerFifoItemAvgRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}
	
	
	@FindBy(xpath = "(//tr[@id='trRender_14'])[1]/td[6]")
	private static WebElement stockLedgerRmaItemReceivedRateAfterSavingStockTransfer;
	
	@FindBy(xpath = "(//tr[@id='trRender_13'])[1]/td[8]")
	private static WebElement stockLedgerRmaItemIssuedRateAfterSavingStockTransfer;
	
	@FindBy(xpath = "(//tr[@id='trRender_14'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemBalanceQtyAfterSavingStockTransfer;
	
	@FindBy(xpath = "(//tr[@id='trRender_14'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemBalanceValueAfterSavingStockTransfer;
	
	@FindBy(xpath = "(//tr[@id='trRender_14'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemAvgRateAfterSavingStockTransfer;
	

	public boolean checkStockLedgerReportForRmaItemAfterStockTransferVoucher()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == true) {
			click(fifoItemCheckBox);

		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == false) {
			click(rmaItemCheckBox);

		}

	
		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);
		

		String actStockLedgerRmaItemReceivedRate 	= getText(stockLedgerRmaItemReceivedRateAfterSavingStockTransfer);
		String expStockLedgerRmaItemReceivedRate	= "834.9155";
		
		String actStockLedgerRmaItemIssuedRate 		= getText(stockLedgerRmaItemIssuedRateAfterSavingStockTransfer);
		String expStockLedgerRmaItemIssuedRate		= "834.9155";		
		

		String actstockLedgerRmaItemBalanceQty		= getText(stockLedgerRmaItemBalanceQtyAfterSavingStockTransfer);
		String expstockLedgerRmaItemBalanceQty 		= "15.0000";

		String actstockLedgerRmaItemBalanceValue 	= getText(stockLedgerRmaItemBalanceValueAfterSavingStockTransfer);
		String expstockLedgerRmaItemBalanceValue 	= "16,927.7910";

		String actstockLedgerrmaItemAvgRate 		= getText(stockLedgerrmaItemAvgRateAfterSavingStockTransfer);
		String expstockLedgerrmaItemAvgRate 		= "1,128.5194";
		
		
		System.out.println("actStockLedgerRmaItemReceivedRate      :" + actStockLedgerRmaItemReceivedRate);
		System.out.println("expStockLedgerRmaItemReceivedRate      :" + expStockLedgerRmaItemReceivedRate);

		System.out.println("actStockLedgerRmaItemIssuedRate        :" + actStockLedgerRmaItemIssuedRate);
		System.out.println("expStockLedgerRmaItemIssuedRate        :" + expStockLedgerRmaItemIssuedRate);

		System.out.println("actstockLedgerRmaItemBalanceQty        :" + actstockLedgerRmaItemBalanceQty);
		System.out.println("expstockLedgerRmaItemBalanceQty        :" + expstockLedgerRmaItemBalanceQty);

		System.out.println("actstockLedgerRmaItemBalanceValue      :" + actstockLedgerRmaItemBalanceValue);
		System.out.println("expstockLedgerRmaItemBalanceValue      :" + expstockLedgerRmaItemBalanceValue);

		System.out.println("actstockLedgerrmaItemAvgRate           :" + actstockLedgerrmaItemAvgRate);
		System.out.println("expstockLedgerrmaItemAvgRate           :" + expstockLedgerrmaItemAvgRate);

		if (actstockLedgerRmaItemBalanceQty.equalsIgnoreCase(expstockLedgerRmaItemBalanceQty)
				&& actstockLedgerRmaItemBalanceValue.equalsIgnoreCase(expstockLedgerRmaItemBalanceValue)
				&& actstockLedgerrmaItemAvgRate.equalsIgnoreCase(expstockLedgerrmaItemAvgRate) && actStockLedgerRmaItemReceivedRate.equalsIgnoreCase(expStockLedgerRmaItemReceivedRate)
				&& actStockLedgerRmaItemIssuedRate.equalsIgnoreCase(expStockLedgerRmaItemIssuedRate)) {

			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return true;
		} else {
			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	public boolean checkSavingStockTransferVocher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
	Thread.sleep(2000);
	click(inventoryMenu);
	click(inventoryTransactionsMenu);
	click(inventoryTransactionsStockMenu);
	click(inventoryTransactionsStockMenuStockTransfersMenu);	
	click(newBtn);
	
	checkValidationMessage("Screen opened");
	
	click(stockTransferDatetxt);
	stockTransferDatetxt.sendKeys(Keys.CONTROL, "a");
	Thread.sleep(2000);
	stockTransferDatetxt.sendKeys(Keys.BACK_SPACE);
	enterText(stockTransferDatetxt, "03/09/2021");
	tab(stockTransferDatetxt);
	tab(stockTransferIssuestxt);
	Thread.sleep(1000);
	enterText(stockTransferWarehousetxt, "SEC");
	Thread.sleep(1000);
	tab(stockTransferWarehousetxt);
	Thread.sleep(2000);
	
	
	click(select1stRow_1stColumn);
	enterText(enterWarehousettxt2, "Test");
	Thread.sleep(2000);
	
	

	tab(enterWarehousettxt2);
	enterText(enterItemttxt, "BR-BATCH");
	Thread.sleep(2000);
	tab(enterItemttxt);
	
	

	
	enterText(enterQuantitytxt, "3");
	tab(enterQuantitytxt);
	enterText(enterRatetxt, "100");
	tab(enterRatetxt);
	Thread.sleep(2000);
	
	
	String actRow1AvgRate				=select1stRow_AvgRate.getText();
	String expRow1AvgRate				="0.0000";
	
	String actRow1OverAllAvgRate		=select1stRow_OverAllAvgRate.getText();
	String expRow1OverAllAvgRate		="0.0000";
	
	
	
	System.out.println("actRow1AvgRate             :"+actRow1AvgRate);
	System.out.println("expRow1AvgRate             :"+expRow1AvgRate);
	
	System.out.println("actRow1OverAllAvgRate      :"+actRow1OverAllAvgRate);
	System.out.println("expRow1OverAllAvgRate      :"+expRow1OverAllAvgRate);
	
	
     click(select1stRow_8thColumn);
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
	
	boolean actBatchpopupIsDisplaying = batchPickOnFIFOIcon.isDisplayed();
	boolean expBatchpopupIsDisplaying = true;
	
	System.out.println("Batch Popup isDisplaying  : "+actBatchpopupIsDisplaying+"  Value Expected  "+expBatchpopupIsDisplaying);
	click(batchPickOnFIFOIcon);
	click(batchOkIcon);
	
	Thread.sleep(2000);
	
	
	String actRow1AvgRate1				=select1stRow_AvgRate.getText();
	String expRow1AvgRate1				="2000.0000";
	
	String actRow1OverAllAvgRate1		=select1stRow_OverAllAvgRate.getText();
	String expRow1OverAllAvgRate1		="2000.0000";
	
	String actRow1BatchText				=select1stRow_BatchText.getText();
	String expRow1BatchText				="01092021PH";
	
	
	
	System.out.println("actRow1AvgRate1             :"+actRow1AvgRate1);
	System.out.println("expRow1AvgRate1             :"+expRow1AvgRate1);
	
	System.out.println("actRow1OverAllAvgRate1      :"+actRow1OverAllAvgRate1);
	System.out.println("expRow1OverAllAvgRate1      :"+expRow1OverAllAvgRate1);
	
	System.out.println("actRow1BatchText            :"+actRow1BatchText);
	System.out.println("expRow1BatchText            :"+expRow1BatchText);



	click(select2ndRow_1stColumn);
	
	enterWarehousettxt2.sendKeys(Keys.CONTROL, "a");
	Thread.sleep(2000);
	enterWarehousettxt2.sendKeys(Keys.BACK_SPACE);	
	enterText(enterWarehousettxt2, "Test");
	Thread.sleep(2000);
	tab(enterWarehousettxt2);
	enterText(enterItemttxt, "WA-RMA");
	Thread.sleep(2000);
	tab(enterItemttxt);
	
	

	
	enterText(enterQuantitytxt, "3");
	tab(enterQuantitytxt);
	enterText(enterRatetxt, "100");
	tab(enterRatetxt);
	Thread.sleep(2000);	
	
	click(select2ndRow_11thColumn);
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

	boolean actRMAPopupIsDisplaying1 = rmaScreenTitle.isDisplayed();
	boolean expRMAPopupIsDisplaying1 = true;

	System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying1 + "  Value Expected  "
			+ expRMAPopupIsDisplaying1);
	
	click(rmaSearchGridField);
	Thread.sleep(1000);
	
	int count = rmaSerialNos.size();
	
	ArrayList<String> SearchRMA = new ArrayList<String>();
	
	for (int i = 0; i < count; i++) 
	{
		String data = rmaSerialNos.get(i).getText();
		SearchRMA.add(data);
		
		if (data.equalsIgnoreCase("01092021PH14") || data.equalsIgnoreCase("01092021PH15") || data.equalsIgnoreCase("01092021PH16")) 
		{
			rmaSerialNoCheckBoxes.get(i).click();
		}
	}
	
	String actSearchRMASerialNoList = SearchRMA.toString();
	String expSearchRMASerialNoList = "[01092021PH14, 01092021PH15, 01092021PH16, 01092021PH17, 01092021PH18, 04082021EXH1]"; 
	
	System.out.println("actSearchRMASerialNoList       :"+actSearchRMASerialNoList);
	System.out.println("expSearchRMASerialNoList       :"+expSearchRMASerialNoList);
	
	
	
	click(searchRmaOkBtn);
	
	int serialcount	=rmaSerialNoList.size();
	ArrayList<String> rmaSerialnos = new ArrayList<String>();
	for (int i = 0; i < serialcount; i++) {
		String data=rmaSerialNoList.get(i).getText();
		rmaSerialnos.add(data);
	}
	
	String actRmaSerialNo 	=rmaSerialnos.toString();
	String expRmaSerialNo	="[01092021PH14, 01092021PH15, 01092021PH16]";
	
	System.out.println("actRmaSerialNo       :"+actRmaSerialNo);
	System.out.println("expRmaSerialNo       :"+expRmaSerialNo);	
	
	
	click(RMAPopupOkBtn);	
	Thread.sleep(1000);
	tab(enterRmatxt);
	
	
	Thread.sleep(3000);
	
	
	String actRow2AvgRate				=select2ndRow_AvgRate.getText();
	String expRow2AvgRate				="1819.0042";
	
	String actRow2OverAllAvgRate		=select2ndRow_OverAllAvgRate.getText();
	String expRow2OverAllAvgRate		="1128.5194";
	
	
	
	System.out.println("actRow2AvgRate             		  :"+actRow2AvgRate);
	System.out.println("expRow2AvgRate             		  :"+expRow2AvgRate);
	
	System.out.println("actRow2OverAllAvgRate             :"+actRow2OverAllAvgRate);
	System.out.println("expRow2OverAllAvgRate             :"+expRow2OverAllAvgRate);
	
		// 3rd Row
	
		click(select3rdRow_1stColumn);
		
		enterWarehousettxt2.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt2.sendKeys(Keys.BACK_SPACE);	
		enterWarehousettxt2.sendKeys(Keys.SPACE);
		Thread.sleep(1000);
		enterText(enterWarehousettxt2, "TEST");
		Thread.sleep(2000);
		tab(enterWarehousettxt2);
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		
		

		
		enterText(enterQuantitytxt, "3");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		Thread.sleep(2000);	
		
		click(select3rdRow_9thColumn);
		Thread.sleep(2000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying + "  Value Expected  "
				+ expBinpopupIsDisplaying);

		click(binSearchBtn);
		Thread.sleep(3000);
		click(binPopup1stRow);
		Thread.sleep(1500);
		click(binPickBtn);
		
		//click(binAutoAllocateBtn);
		Thread.sleep(2000);
		click(binOkBtn);
		
		Thread.sleep(2000);
		
		tab(bin1Txt);
		
		//bin2Txt.sendKeys(Keys.SPACE);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying1 = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying1 = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying1 + "  Value Expected  "
				+ expBinpopupIsDisplaying1);

		click(binSearchBtn);
		Thread.sleep(2000);
		click(binPopup1stRow);
		Thread.sleep(1500);
		click(binPickBtn);
		//click(binAutoAllocateBtn);
		Thread.sleep(2000);
		click(binOkBtn);

	
	
			

	Thread.sleep(2000);
	click(voucherSaveBtn);
	Thread.sleep(2000);

	HashSet<String> actMessage = new HashSet<String>();

	for (int i = 0; i < 2; i++) {
		String data = checkValidationMessage("");
		actMessage.add(data);
	}

	HashSet<String> expMessage = new HashSet<String>();

	expMessage.add("Voucher saved successfully : 71");
	expMessage.add("Saving in background.");

	System.out.println("Actual Message    : " + actMessage);
	System.out.println("Expected Message  : " + expMessage);

	if(actMessage.equals(expMessage) && actRow1AvgRate.equalsIgnoreCase(expRow1AvgRate) && actRMAPopupIsDisplaying1==expRMAPopupIsDisplaying1 
			&& actRow1OverAllAvgRate.equalsIgnoreCase(expRow1OverAllAvgRate) && actRow1AvgRate1.equalsIgnoreCase(expRow1AvgRate1)
			&& actRow1OverAllAvgRate1.equalsIgnoreCase(expRow1OverAllAvgRate1) && actRow1BatchText.equalsIgnoreCase(expRow1BatchText)
			&& actSearchRMASerialNoList.equalsIgnoreCase(expSearchRMASerialNoList) && actRmaSerialNo.equalsIgnoreCase(expRmaSerialNo)
			&& actRow2AvgRate.equalsIgnoreCase(expRow2AvgRate) && actRow2OverAllAvgRate.equalsIgnoreCase(expRow2OverAllAvgRate)){	
		click(new_CloseBtn);			
		Thread.sleep(2000);
		click(voucherhomeCloseBtn);
		return true;
		
	} else {
		click(new_CloseBtn);			
		Thread.sleep(2000);
		click(voucherhomeCloseBtn);
		return false;

	  }
	}
	
	public boolean validteItemQueryStocksInBatchItemAfterSavingStockTransferVoucher2() throws InterruptedException {

		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(itemqueryMenu);
		click(itemTxt);
		Thread.sleep(2000);
		enterText(itemTxt, "BR-BATCH");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "15.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "1,267.7481";

		String actlstockValueTxt 	 = getText(stockValueTxt);
		String expstockValueTxt		 = "19,016.2222";

		System.out.println("=====Stocks In Batch Items======");

		System.out.println("actlOpenStockQuantity      :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity       :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock           :" + actlCurrentStock);
		System.out.println("expCurrentStock            :" + expCurrentStock);

		System.out.println("actlavgStockRate           :" + actlavgStockRate);
		System.out.println("expavgStockRate            :" + expavgStockRate);

		System.out.println("actlstockValueTxt          :" + actlstockValueTxt);
		System.out.println("expstockValueTxt           :" + expstockValueTxt);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)) {
			return true;
		} else {
			return false;
		}

	}

	public boolean validteItemQueryStocksInFifoItemAfterSavingStockTransferVoucher2() throws InterruptedException {

		click(itemTxt);
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.BACK_SPACE);
		enterText(itemTxt, "FIFO-BIN");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(2000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "15.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "1,800.0000";
 
		String actlstockValueTxt     = getText(stockValueTxt);
		String expstockValueTxt      = "27,000.0000";

		System.out.println("=====Stocks In FIFO Items======");

		System.out.println("actlOpenStockQuantity :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity  :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock      :" + actlCurrentStock);
		System.out.println("expCurrentStock       :" + expCurrentStock);

		System.out.println("actlavgStockRate      :" + actlavgStockRate);
		System.out.println("expavgStockRate       :" + expavgStockRate);

		System.out.println("actlstockValueTxt     :" + actlstockValueTxt);
		System.out.println("expstockValueTxt      :" + expstockValueTxt);

		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)) {
			return true;
		} else {
			return false;
		}

	}

	public boolean validteItemQueryStocksInWAItemAfterSavingStockTransferVoucher2() throws InterruptedException {

		click(itemTxt);
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		itemTxt.sendKeys(Keys.BACK_SPACE);
		enterText(itemTxt, "WA-RMA");
		Thread.sleep(3000);
		click(getStockBtn);
		Thread.sleep(3000);

		String actlOpenStockQuantity = getText(openStockQuantityTxt);
		String expOpenStockQuantity  = "20.000";

		String actlCurrentStock      = getText(currentStockTxt);
		String expCurrentStock       = "15.000";

		String actlavgStockRate      = getText(avgStockRateTxt);
		String expavgStockRate       = "1,266.6164";

		String actlstockValueTxt     = getText(stockValueTxt);
		String expstockValueTxt      = "18,999.2453";

		ScrollToElement(balanceByWarehouseBtn);
		click(balanceByWarehouseBtn);

		Thread.sleep(3000);

		int count = row1BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow1 = new ArrayList<String>();

		for (int i = 0; i < count; i++) {
			String data = row1BalanceByWareHouseList.get(i).getText();
			warehouseRow1.add(data);
		}

		String actwarehouseRow1List = warehouseRow1.toString();
		String expWarehouseRow1List = "[1, HYD, 4.000, 0.000, 4.000, 834.9155, 3,339.6618]";

		int count1 = row2BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow2 = new ArrayList<String>();

		for (int i = 0; i < count1; i++) {
			String data = row2BalanceByWareHouseList.get(i).getText();
			warehouseRow2.add(data);
		}

		String actwarehouseRow2List = warehouseRow2.toString();
		String expWarehouseRow2List = "[2, SEC, 3.000, 0.000, 3.000, 1,819.0042, 5,457.0125]";

		int count2 = row3BalanceByWareHouseList.size();

		ArrayList<String> warehouseRow3 = new ArrayList<String>();

		for (int i = 0; i < count2; i++) {
			String data = row3BalanceByWareHouseList.get(i).getText();
			warehouseRow3.add(data);
		}

		String actwarehouseRow3List = warehouseRow3.toString();
		String expWarehouseRow3List = "[3, TEST, 8.000, 0.000, 8.000, 998.0464, 7,984.3713]";
		
		


		System.out.println("=====Stocks In WA Items======");

		System.out.println("actlOpenStockQuantity :" + actlOpenStockQuantity);
		System.out.println("expOpenStockQuantity  :" + expOpenStockQuantity);

		System.out.println("actlCurrentStock      :" + actlCurrentStock);
		System.out.println("expCurrentStock       :" + expCurrentStock);

		System.out.println("actlavgStockRate      :" + actlavgStockRate);
		System.out.println("expavgStockRate       :" + expavgStockRate);

		System.out.println("actlstockValueTxt     :" + actlstockValueTxt);
		System.out.println("expstockValueTxt      :" + expstockValueTxt);

		System.out.println("===========Balance By Warehouse=======");

		System.out.println(actwarehouseRow1List);
		System.out.println(expWarehouseRow1List);

		System.out.println(actwarehouseRow2List);
		System.out.println(expWarehouseRow2List);

		System.out.println(actwarehouseRow3List);
		System.out.println(expWarehouseRow3List);

		
	
		
		
		if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
				&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
				&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
				&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)
				&& actwarehouseRow1List.equalsIgnoreCase(expWarehouseRow1List)
				&& actwarehouseRow2List.equalsIgnoreCase(expWarehouseRow2List)
				&& actwarehouseRow3List.equalsIgnoreCase(expWarehouseRow3List)) 
		{
			return true;
		} else {
			return false;
		}

	}
	
	@FindBy(xpath = "(//tr[@id='trRender_14'])[1]/td[6]")
	private static WebElement stockLedgerBrBatchReceivedRateAfterSavingStockTransfer2;
	
	@FindBy(xpath = "(//tr[@id='trRender_13'])[1]/td[8]")
	private static WebElement stockLedgerBrBatchIssuedRateAfterSavingStockTransfer2;	
	
	@FindBy(xpath = "(//tr[@id='trRender_14'])[1]/td[13]")
	private static WebElement stockLedgerBrBatchAvgRateAfterSavingStockTransfer2;
	
	@FindBy(xpath = "(//tr[@id='trRender_14'])[1]/td[9]")
	private static WebElement stockLedgerBrBatchBalanceQtyAfterSavingStockTransfer2;
	
	@FindBy(xpath = "(//tr[@id='trRender_14'])[1]/td[12]")
	private static WebElement stockLedgerBrBatchBalanceValueAfterSavingStockTransfer2;
	
	
	public boolean checkStockLedgerReportForBrBatchItemAfterStockTransferVoucher2()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryStockLedgerMenu);
		Thread.sleep(3000);

		int rowcount = stockLedgerHometableRowCount.size();

		System.out.println(rowcount);

		for (int i = 0; i < rowcount; i++) {
			String actName = ledgerHometableItemNamesList.get(i).getText();

			System.out.println(actName);

			if (actName.equalsIgnoreCase("BR-BATCH")) {
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);

		String actStockLedgerBrBatchReceivedRate 	= getText(stockLedgerBrBatchReceivedRateAfterSavingStockTransfer2);
		String expStockLedgerBrBatchReceivedRate	= "2,000.0000";
		
		String actStockLedgerBrBatchIssuedRate 		= getText(stockLedgerBrBatchIssuedRateAfterSavingStockTransfer2);
		String expStockLedgerBrBatchIssuedRate		= "2,000.0000";
		
		
		String actStockLedgerBrBatchBalanceQty 		= getText(stockLedgerBrBatchBalanceQtyAfterSavingStockTransfer2);
		String expStockLedgerBrBatchBalanceQty 	    = "15.0000";

		String actStockLedgerBrBatchBalanceValue    = getText(stockLedgerBrBatchBalanceValueAfterSavingStockTransfer2);
		String expStockLedgerBrBatchBalanceValue    = "19,016.2222";

		String actStockLedgerBrBatchAvgRate 		= getText(stockLedgerBrBatchAvgRateAfterSavingStockTransfer2);
		String expStockLedgerBrBatchAvgRate			= "1,267.7481";

		
		
		System.out.println("actStockLedgerBrBatchReceivedRate    :" + actStockLedgerBrBatchReceivedRate);
		System.out.println("expStockLedgerBrBatchReceivedRate    :" + expStockLedgerBrBatchReceivedRate);

		System.out.println("actStockLedgerBrBatchIssuedRate      :" + actStockLedgerBrBatchIssuedRate);
		System.out.println("expStockLedgerBrBatchIssuedRate      :" + expStockLedgerBrBatchIssuedRate);
		
		System.out.println("actStockLedgerBrBatchBalanceQty      :" + actStockLedgerBrBatchBalanceQty);
		System.out.println("expStockLedgerBrBatchBalanceQty      :" + expStockLedgerBrBatchBalanceQty);


		System.out.println("actStockLedgerBrBatchBalanceValue    :" + actStockLedgerBrBatchBalanceValue);
		System.out.println("expStockLedgerBrBatchBalanceValue    :" + expStockLedgerBrBatchBalanceValue);

		System.out.println("actStockLedgerBrBatchAvgRate         :" + actStockLedgerBrBatchAvgRate);
		System.out.println("expStockLedgerBrBatchAvgRate         :" + expStockLedgerBrBatchAvgRate);

		if (actStockLedgerBrBatchBalanceQty.equalsIgnoreCase(expStockLedgerBrBatchBalanceQty)
				&& actStockLedgerBrBatchBalanceValue.equalsIgnoreCase(expStockLedgerBrBatchBalanceValue)
				&& actStockLedgerBrBatchAvgRate.equalsIgnoreCase(expStockLedgerBrBatchAvgRate) && actStockLedgerBrBatchReceivedRate.equalsIgnoreCase(expStockLedgerBrBatchReceivedRate)
				&& actStockLedgerBrBatchIssuedRate.equalsIgnoreCase(expStockLedgerBrBatchIssuedRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}

	 
	@FindBy(xpath = "(//tr[@id='trRender_15'])[1]/td[13]")
	private static WebElement stockLedgerFifoItemAvgRateAfterSavingStockTransfer2;
	
	@FindBy(xpath = "(//tr[@id='trRender_15'])[1]/td[9]")
	private static WebElement stockLedgerFifoItemBalanceQtyAfterSavingStockTransfer2;
	
	@FindBy(xpath = "(//tr[@id='trRender_15'])[1]/td[12]")
	private static WebElement stockLedgerFifoItemBalanceValueAfterSavingStockTransfer2;
	
	
	
	public boolean checkStockLedgerReportForFifoItemAfterStockTransferVoucher2()  
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == false) {
			click(fifoItemCheckBox);
		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == true) {
			click(rmaItemCheckBox);

		}

		
		dropDown(sl_DateOptionDropdown, "1");	
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);

		String actstockLedgerFifoItemBalanceQty   = getText(stockLedgerFifoItemBalanceQtyAfterSavingStockTransfer2);
		String expstockLedgerFifoItemBalanceQty   = "15.0000";

		String actstockLedgerFifoItemBalanceValue = getText(stockLedgerFifoItemBalanceValueAfterSavingStockTransfer2);
		String expstockLedgerFifoItemBalanceValue = "27,000.0000";

		String actstockLedgerFifoItemAvgRate      = getText(stockLedgerFifoItemAvgRateAfterSavingStockTransfer2);
		String expstockLedgerFifoItemAvgRate      = "1,800.0000";

		System.out.println("actstockLedgerFifoItemBalanceQty      :" + actstockLedgerFifoItemBalanceQty);
		System.out.println("expstockLedgerFifoItemBalanceQty      :" + expstockLedgerFifoItemBalanceQty);

		System.out.println("actstockLedgerFifoItemBalanceValue    :" + actstockLedgerFifoItemBalanceValue);
		System.out.println("expstockLedgerFifoItemBalanceValue    :" + expstockLedgerFifoItemBalanceValue);

		System.out.println("actstockLedgerFifoItemAvgRate         :" + actstockLedgerFifoItemAvgRate);
		System.out.println("expstockLedgerFifoItemAvgRate         :" + expstockLedgerFifoItemAvgRate);

		if (actstockLedgerFifoItemBalanceQty.equalsIgnoreCase(expstockLedgerFifoItemBalanceQty)
				&& actstockLedgerFifoItemBalanceValue.equalsIgnoreCase(expstockLedgerFifoItemBalanceValue)
				&& actstockLedgerFifoItemAvgRate.equalsIgnoreCase(expstockLedgerFifoItemAvgRate)) {

			click(report_CloseBtn);

			return true;
		} else {
			click(report_CloseBtn);

			return false;
		}
	}
	
	
	@FindBy(xpath = "(//tr[@id='trRender_16'])[1]/td[6]")
	private static WebElement stockLedgerRmaItemReceivedRateAfterSavingStockTransfer2;
	
	@FindBy(xpath = "(//tr[@id='trRender_15'])[1]/td[8]")
	private static WebElement stockLedgerRmaItemIssuedRateAfterSavingStockTransfer2;
	
	@FindBy(xpath = "(//tr[@id='trRender_16'])[1]/td[9]")
	private static WebElement stockLedgerRmaItemBalanceQtyAfterSavingStockTransfer2;
	
	@FindBy(xpath = "(//tr[@id='trRender_16'])[1]/td[12]")
	private static WebElement stockLedgerRmaItemBalanceValueAfterSavingStockTransfer2;
	
	@FindBy(xpath = "(//tr[@id='trRender_16'])[1]/td[13]")
	private static WebElement stockLedgerrmaItemAvgRateAfterSavingStockTransfer2;
	

	public boolean checkStockLedgerReportForRmaItemAfterStockTransferVoucher2()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);

		if (fifoItemCheckBox.isSelected() == true) {
			click(fifoItemCheckBox);

		}

		if (brBatchItemCheckBox.isSelected() == true) {
			click(brBatchItemCheckBox);
		}

		if (rmaItemCheckBox.isSelected() == false) {
			click(rmaItemCheckBox);

		}

	
		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(2000);
		

		String actStockLedgerRmaItemReceivedRate 	= getText(stockLedgerRmaItemReceivedRateAfterSavingStockTransfer2);
		String expStockLedgerRmaItemReceivedRate	= "1,819.0042";
		
		String actStockLedgerRmaItemIssuedRate 		= getText(stockLedgerRmaItemIssuedRateAfterSavingStockTransfer2);
		String expStockLedgerRmaItemIssuedRate		= "1,819.0042";		
		

		String actstockLedgerRmaItemBalanceQty		= getText(stockLedgerRmaItemBalanceQtyAfterSavingStockTransfer2);
		String expstockLedgerRmaItemBalanceQty 		= "15.0000";

		String actstockLedgerRmaItemBalanceValue 	= getText(stockLedgerRmaItemBalanceValueAfterSavingStockTransfer2);
		String expstockLedgerRmaItemBalanceValue 	= "18,999.2453";

		String actstockLedgerrmaItemAvgRate 		= getText(stockLedgerrmaItemAvgRateAfterSavingStockTransfer2);
		String expstockLedgerrmaItemAvgRate 		= "1,266.6164";
		
		
		System.out.println("actStockLedgerRmaItemReceivedRate    :" + actStockLedgerRmaItemReceivedRate);
		System.out.println("expStockLedgerRmaItemReceivedRate    :" + expStockLedgerRmaItemReceivedRate);

		System.out.println("actStockLedgerRmaItemIssuedRate      :" + actStockLedgerRmaItemIssuedRate);
		System.out.println("expStockLedgerRmaItemIssuedRate      :" + expStockLedgerRmaItemIssuedRate);

		System.out.println("actstockLedgerRmaItemBalanceQty      :" + actstockLedgerRmaItemBalanceQty);
		System.out.println("expstockLedgerRmaItemBalanceQty      :" + expstockLedgerRmaItemBalanceQty);

		System.out.println("actstockLedgerRmaItemBalanceValue    :" + actstockLedgerRmaItemBalanceValue);
		System.out.println("expstockLedgerRmaItemBalanceValue    :" + expstockLedgerRmaItemBalanceValue);

		System.out.println("actstockLedgerrmaItemAvgRate         :" + actstockLedgerrmaItemAvgRate);
		System.out.println("expstockLedgerrmaItemAvgRate         :" + expstockLedgerrmaItemAvgRate);

		if (actstockLedgerRmaItemBalanceQty.equalsIgnoreCase(expstockLedgerRmaItemBalanceQty)
				&& actstockLedgerRmaItemBalanceValue.equalsIgnoreCase(expstockLedgerRmaItemBalanceValue)
				&& actstockLedgerrmaItemAvgRate.equalsIgnoreCase(expstockLedgerrmaItemAvgRate) && actStockLedgerRmaItemReceivedRate.equalsIgnoreCase(expStockLedgerRmaItemReceivedRate)
				&& actStockLedgerRmaItemIssuedRate.equalsIgnoreCase(expStockLedgerRmaItemIssuedRate)) {

			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return true;
		} else {
			click(report_CloseBtn);
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	@FindBy(xpath = "//a[@id='2035']/span")
	private static WebElement inventoryTransactionsExcessesStocks;
	
	@FindBy(xpath="//input[@id='id_header_2']")
	private static WebElement  excessesStockDatetxt;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[10]")
	private static WebElement  select4thRow_9thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[11]")
	private static WebElement  select5thRow_10thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[11]")
	private static WebElement  select6thRow_10thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[7]")
	private static WebElement  select6thRow_ExcessStockOverAllAvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[7]")
	private static WebElement  select5thRow_ExcessStockOverAllAvgRate;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[9]")
	private static WebElement  select2ndRow_8thColumn;
	
	
	
	public boolean  checkSavingexcessInStocksVoucherWithALLItemsInInvUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{

		Thread.sleep(3000);
		click(userNameDisplay);
		Thread.sleep(1000);
		click(logoutOption);		
		Thread.sleep(3000);
		
		String unamelt	="inv";

		String pawslt	="su";
		
		LoginPage.enterUserName(unamelt);	     
        LoginPage.enterPassword(pawslt);
        
		
		 Select company = new Select(companyDropDownList);
		 company.selectByValue("108");
		 Thread.sleep(2000);
		 
  
        LoginPage.clickOnSignInBtn();  
        
        Thread.sleep(3000);
        
        System.out.println("Signing");
        
        // Login User Name is Verify
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		String actUserInfo                      = userNameDisplay.getText();
		String expUserInfo                      = "INV";	

		System.out.println("Login With Valid User and Check User Name Display Value Actual : ." + actUserInfo + ". Value Expected : " + expUserInfo);
		
		
		// Login Company Name is Verify
		Thread.sleep(1000);
		click(companyLogo);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
		String actGetLoginCompanyNameInformation = companyName.getText();
		String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
		String expGetLoginCompanyName            = "Maintain Quantity a";   
		companyLogo.click();
		
		System.out.println("Company Name Display Value Actual                : ." + actGetLoginCompanyName +  ". Value Expected : " + expGetLoginCompanyName);
		
		Thread.sleep(2000);
		click(inventoryMenu);
		click(inventoryTransactionsMenu);
		click(inventoryTransactionsStockMenu);
		click(inventoryTransactionsExcessesStocks);
		
		Thread.sleep(3000);
		
		click(newBtn);
		
		checkValidationMessage("Screen opened");
		
		click(excessesStockDatetxt);
		excessesStockDatetxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		excessesStockDatetxt.sendKeys(Keys.BACK_SPACE);
		enterText(excessesStockDatetxt, "04/09/2021");
		tab(excessesStockDatetxt);
		
		
		click(select1stRow_1stColumn);
		
		
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(1000);
		tab(enterWarehousettxt);
				
		enterText(enterItemttxt, "BR-BATCH");
		Thread.sleep(2000);
		tab(enterItemttxt);
		
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		
		click(select1stRow_8thColumn);
		
		enterText(enterBatchtxt, "04092021EXH");
		tab(enterBatchtxt);
		
		// 2nd Row
		
		click(select2ndRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);		
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(1000);
		tab(enterWarehousettxt);				
		enterText(enterItemttxt, "BR-BATCH");
		Thread.sleep(2000);
		tab(enterItemttxt);		
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "200");
		tab(enterRatetxt);
		
		
		click(select2ndRow_8thColumn);
		
		enterText(enterBatchtxt, "04092021EXS");
		tab(enterBatchtxt);
		
		//3rd Row
		
		click(select3rdRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		
		
		click(select3rdRow_9thColumn);		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying + "  Value Expected  "
				+ expBinpopupIsDisplaying);

		click(binSearchBtn);
		Thread.sleep(3000);
		click(binPopup1stRow);
		Thread.sleep(1500);
		click(binPickBtn);
		//click(binAutoAllocateBtn);
		Thread.sleep(2000);
		click(binOkBtn);
		Thread.sleep(2000);

		// 4th Row

		click(select4thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		
		
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "200");
		tab(enterRatetxt);
		
		
		click(select4thRow_9thColumn);		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying1 = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying1 = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying1 + "  Value Expected  "
				+ expBinpopupIsDisplaying1);

		click(binSearchBtn);
		Thread.sleep(3000);
		click(binPopup1stRow);
		Thread.sleep(1500);
		click(binPickBtn);
		//click(binAutoAllocateBtn);
		Thread.sleep(2000);
		click(binOkBtn);

		// 5th Row

		click(select5thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);
		tab(enterItemttxt);
		
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		
		click(select5thRow_10thColumn);
		
		Thread.sleep(2000);
		
		String actual5throwAvgRate 			= select5thRow_AvgRate.getText();
		String exp5throwAvgRate			    = "834.9155";
		String actual5throwOverallAvgRate   = select5thRow_ExcessStockOverAllAvgRate.getText();
		String exp5throwoverallAvgRate      = "1305.5580";
		
		System.out.println("actual5throwAvgRate       	 :" + actual5throwAvgRate);
		System.out.println("exp5throwAvgRate             :" + exp5throwAvgRate);

		System.out.println("actual5throwOverallAvgRate   :" + actual5throwOverallAvgRate);
		System.out.println("exp5throwoverallAvgRate      :" + exp5throwoverallAvgRate);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying + "  Value Expected  "
				+ expRMAPopupIsDisplaying);
		
		enterText(rmaSerialNumberTxtField, "04092021EXH1,1");
		click(rmaAddBtn);
		Thread.sleep(2000);
		click(RMAPopupOkBtn);
		Thread.sleep(3000);
		tab(enterRmatxt);

		// 6th Row
		Thread.sleep(1000);
		click(select6thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);		
		tab(enterItemttxt);
		
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "200");
		tab(enterRatetxt);
		
		click(select6thRow_10thColumn);
		
		Thread.sleep(2000);
		
		String actual6throwAvgRate 				= select6thRow_AvgRate.getText();
		String exp6throwAvgRate 				= "1819.0042";
		String actual6throwOverallAvgRate		= select6thRow_ExcessStockOverAllAvgRate.getText();
		String exp6throwoverallAvgRate			= "1305.5580";
		
		System.out.println("actual6throwAvgRate      	 :" + actual6throwAvgRate);
		System.out.println("exp6throwAvgRate          	 :" + exp6throwAvgRate);

		System.out.println("actual6throwOverallAvgRate   :" + actual6throwOverallAvgRate);
		System.out.println("exp6throwoverallAvgRate      :" + exp6throwoverallAvgRate);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying1 = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying1 = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying1 + "  Value Expected  "
				+ expRMAPopupIsDisplaying1);
		enterText(rmaSerialNumberTxtField, "04092021EXS1,1");
		click(rmaAddBtn);
		Thread.sleep(2000);
		click(RMAPopupOkBtn);
		Thread.sleep(2000);		
		tab(enterRmatxt);
		
		
		Thread.sleep(3000);
		click(voucherSaveBtn);

		HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) {
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : 33");
		expMessage.add("Saving in background.");

		System.out.println("Actual Message         : " + actMessage);
		System.out.println("Expected Message       : " + expMessage);

		if (actMessage.equals(expMessage) && actual5throwAvgRate.equalsIgnoreCase(exp5throwAvgRate)
				&& actual5throwOverallAvgRate.equalsIgnoreCase(exp5throwoverallAvgRate)
				&& actual6throwAvgRate.equalsIgnoreCase(exp6throwAvgRate)
				&& actual6throwOverallAvgRate.equalsIgnoreCase(exp6throwoverallAvgRate)) {
			
			click(new_CloseBtn);			
			Thread.sleep(1000);
			click(voucherhomeCloseBtn);
			return true;
			
		} else {
			click(new_CloseBtn);			
			Thread.sleep(1000);
			click(voucherhomeCloseBtn);
			return false;

		}

	}
	
	@FindBy(xpath = "//a[@id='2034']/span")
	private static WebElement inventoryTransactionsShortagesInStocks;
	
	@FindBy(xpath="//input[@id='id_header_2']")
	private static WebElement shortagesInStockDatetxt;
	
	@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_body']/tr[1]/td[2]")
	private static WebElement  batchPopUPRow1BatchNoTxt;
	
	@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_body']/tr[2]/td[2]")
	private static WebElement  batchPopUPRow2BatchNoTxt;
	
	@FindBy(xpath = "//*[@id='Id_MRPBatch_Grid_body']/tr[3]/td[2]")
	private static WebElement batchPopUPRow3BatchNoTxt;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td")
	private static List<WebElement> Row3ShortageInStockVoucherBodyList;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td")
	private static List<WebElement> Row4ShortageInStockVoucherBodyList;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td")
	private static List<WebElement> Row5ShortageInStockVoucherBodyList;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td")
	private static List<WebElement> Row6ShortageInStockVoucherBodyList;
	
	public boolean  checkSavingShortagesinStockVoucherWithBrItemHydSecWareHosueInSUUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{
		
		
		
		Thread.sleep(3000);
		click(userNameDisplay);
		Thread.sleep(1000);
		click(logoutOption);		
		Thread.sleep(3000);
		
		String unamelt ="su";

		String pawslt  ="su";
		
		LoginPage.enterUserName(unamelt);	     
        LoginPage.enterPassword(pawslt);
        
		
		 Select company = new Select(companyDropDownList);
		 company.selectByValue("108");
		 Thread.sleep(2000);
		 
  
        LoginPage.clickOnSignInBtn();  
        
        Thread.sleep(3000);
        
        System.out.println("Signing");
        
        // Login User Name is Verify
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		String actUserInfo                      = userNameDisplay.getText();
		String expUserInfo                      = "SU";	

		System.out.println("Login With Valid User and Check User Name Display Value Actual : ." + actUserInfo + ". Value Expected : " + expUserInfo);
		
		
		// Login Company Name is Verify
		Thread.sleep(1000);
		click(companyLogo);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
		String actGetLoginCompanyNameInformation = companyName.getText();
		String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
		String expGetLoginCompanyName            = "Maintain Quantity a";   
		companyLogo.click();
		
		System.out.println("Company Name Display Value Actual                : ." + actGetLoginCompanyName +  ". Value Expected : " + expGetLoginCompanyName);
		
		Thread.sleep(2000);

		
		
		click(inventoryMenu);
		click(inventoryTransactionsMenu);
		click(inventoryTransactionsStockMenu);
		click(inventoryTransactionsShortagesInStocks);
		click(newBtn);	
		
		checkValidationMessage("Screen opened");
		
		
		click(shortagesInStockDatetxt);
		shortagesInStockDatetxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		shortagesInStockDatetxt.sendKeys(Keys.BACK_SPACE);
		enterText(shortagesInStockDatetxt, "05/09/2021");
		tab(shortagesInStockDatetxt);
		
		
		click(select1stRow_1stColumn);		
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(1000);
		tab(enterWarehousettxt);	
		
		enterText(enterItemttxt, "BR-BATCH");
		Thread.sleep(2000);
		tab(enterItemttxt);	
		
		enterText(enterQuantitytxt, "3");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);	
		
		click(select1stRow_8thColumn);
		
		Thread.sleep(2000);
		
		String actRow1BatchNo		=batchPopUPRow1BatchNoTxt.getText();
		String expRow1BatchNo		="01092021PH";
		
		String actRow2BatchNo		=batchPopUPRow2BatchNoTxt.getText();
		String expRow2BatchNo		="04092021EXH";
		
		String actRow1Baseqty		=batchPopUPRow1BaseQuantityTxt.getText();
		String expRow1Baseqty		="4.000";
		
		String actRow1Rate			=batchPopUPRow1RateTxt.getText();
		String expRow1Rate			="1000";
		
		String actRow2Baseqty		=batchPopUPRow2BaseQuantityTxt.getText();
		String expRow2Baseqty		="1.000";
		
		String actRow2Rate			=batchPopUPRow2RateTxt.getText();
		String expRow2Rate			="100";
		
	
		System.out.println("actRow1BatchNo             :"+actRow1BatchNo);
		System.out.println("expRow1BatchNo             :"+expRow1BatchNo);
		
		System.out.println("actRow2BatchNo             :"+actRow2BatchNo);
		System.out.println("expRow2BatchNo             :"+expRow2BatchNo);		
		
		
		System.out.println("actRow1Baseqty             :"+actRow1Baseqty);
		System.out.println("expRow1Baseqty             :"+expRow1Baseqty);
		
		System.out.println("actRow1Rate                :"+actRow1Rate);
		System.out.println("expRow1Rate                :"+expRow1Rate);
		
		System.out.println("actRow2Baseqty             :"+actRow2Baseqty);
		System.out.println("expRow2Baseqty             :"+expRow2Baseqty);
		
		System.out.println("actRow2Rate                :"+actRow2Rate);
		System.out.println("expRow2Rate                :"+expRow2Rate);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		
		boolean actBatchpopupIsDisplaying = batchPickOnFIFOIcon.isDisplayed();
		boolean expBatchpopupIsDisplaying = true;
		
		System.out.println("Batch Popup isDisplaying  : " + actBatchpopupIsDisplaying + "  Value Expected  "
				+ expBatchpopupIsDisplaying);
		click(batchPickOnFIFOIcon);
		click(batchOkIcon);

		// 2nd Row

		click(select2ndRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(3000);
		tab(enterWarehousettxt);
		
		enterText(enterItemttxt, "BR-BATCH");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);	
		
		enterText(enterQuantitytxt, "3");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);	
		
		click(select2ndRow_8thColumn);
		
		Thread.sleep(2000);
		
		String actRow1BatchNoSecLocation		=batchPopUPRow1BatchNoTxt.getText();
		String expRow1BatchNoSecLocation		="01092021PH";
		
		String actRow2BatchNoSecLocation		=batchPopUPRow2BatchNoTxt.getText();
		String expRow2BatchNoSecLocation		="04092021EXS";
		
		String actRow1BaseqtySecLocation		=batchPopUPRow1BaseQuantityTxt.getText();
		String expRow1BaseqtySecLocation		="3.000";
		
		String actRow1RateSecLocation			=batchPopUPRow1RateTxt.getText();
		String expRow1RateSecLocation			="2000";
		
		String actRow2BaseqtySecLocation		=batchPopUPRow2BaseQuantityTxt.getText();
		String expRow2BaseqtySecLocation		="1.000";
		
		String actRow2RateSecLocation			=batchPopUPRow2RateTxt.getText();
		String expRow2RateSecLocation			="200";
		
	
		System.out.println("actRow1BatchNoSecLocation             :"+actRow1BatchNoSecLocation);
		System.out.println("expRow1BatchNoSecLocation             :"+expRow1BatchNoSecLocation);
		
		System.out.println("actRow2BatchNoSecLocation             :"+actRow2BatchNoSecLocation);
		System.out.println("expRow2BatchNoSecLocation             :"+expRow2BatchNoSecLocation);
		
		
		
		System.out.println("actRow1BaseqtySecLocation             :"+actRow1BaseqtySecLocation);
		System.out.println("expRow1BaseqtySecLocation             :"+expRow1BaseqtySecLocation);
		
		System.out.println("actRow1RateSecLocation                :"+actRow1RateSecLocation);
		System.out.println("expRow1RateSecLocation                :"+expRow1RateSecLocation);
		
		System.out.println("actRow2BaseqtySecLocation             :"+actRow2BaseqtySecLocation);
		System.out.println("expRow2BaseqtySecLocation             :"+expRow2BaseqtySecLocation);
		
		System.out.println("actRow2RateSecLocation                :"+actRow2RateSecLocation);
		System.out.println("expRow2RateSecLocation                :"+expRow2RateSecLocation);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		
		boolean actBatchpopupIsDisplaying1 = batchPickOnFIFOIcon.isDisplayed();
		boolean expBatchpopupIsDisplaying1 = true;
		
		System.out.println("Batch Popup isDisplaying  : " + actBatchpopupIsDisplaying1 + "  Value Expected  "
				+ expBatchpopupIsDisplaying1);
		click(batchPickOnFIFOIcon);
		click(batchOkIcon);
		
		if(actRow1BatchNo.equalsIgnoreCase(expRow1BatchNo) && actRow2BatchNo.equalsIgnoreCase(expRow2BatchNo) 
				&& actRow1Baseqty.equalsIgnoreCase(expRow1Baseqty) && actRow1Rate.equalsIgnoreCase(expRow1Rate) 
				&& actRow2Baseqty.equalsIgnoreCase(expRow2Baseqty) && actRow2Rate.equalsIgnoreCase(expRow2Rate)
				&& actRow1BatchNoSecLocation.equalsIgnoreCase(expRow1BatchNoSecLocation) && actRow2BatchNoSecLocation.equalsIgnoreCase(expRow2BatchNoSecLocation)
				&& actRow1BaseqtySecLocation.equalsIgnoreCase(expRow1BaseqtySecLocation) && actRow1RateSecLocation.equalsIgnoreCase(expRow1RateSecLocation)
				&& actRow2BaseqtySecLocation.equalsIgnoreCase(expRow2BaseqtySecLocation) && actRow2RateSecLocation.equalsIgnoreCase(expRow2RateSecLocation)){
			
			return true;
		}
			else{
				return false;
			}
		
	}
	
	public boolean  checkSavingShortagesinStockVoucherWithBrItemTestWareHosue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{		
		
		
		// 3rd Row

				click(select3rdRow_1stColumn);
				Thread.sleep(2000);
				enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
				Thread.sleep(2000);
				enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
				enterText(enterWarehousettxt, "TEST");
				Thread.sleep(3000);
				tab(enterWarehousettxt);
				
				enterText(enterItemttxt, "BR-BATCH");
				Thread.sleep(2000);
				tab(enterItemttxt);
				Thread.sleep(2000);	
				
				enterText(enterQuantitytxt, "6");
				tab(enterQuantitytxt);
				
				enterText(enterRatetxt, "200");
				tab(enterRatetxt);	
				
				click(select3rdRow_8thColumn);
				
				Thread.sleep(3000);
				
				String actRow1BatchNoTestLocation		=batchPopUPRow1BatchNoTxt.getText();
				String expRow1BatchNoTestLocation		="01082021PS";
				
				String actRow2BatchNoTestLocation		=batchPopUPRow2BatchNoTxt.getText();
				String expRow2BatchNoTestLocation		="01092021PH";
				
				String actRow3BatchNoTestLocation		=batchPopUPRow3BatchNoTxt.getText();
				String expRow3BatchNoTestLocation		="01092021PH";
				
				String actRow1BaseqtyTestLocation		=batchPopUPRow1BaseQuantityTxt.getText();
				String expRow1BaseqtyTestLocation		="2.000";
				
				String actRow1RateTestLocation			=batchPopUPRow1RateTxt.getText();
				String expRow1RateTestLocation			="8.1111";
				
				String actRow2BaseqtyTestLocation		=batchPopUPRow2BaseQuantityTxt.getText();
				String expRow2BaseqtyTestLocation		="3.000";
				
				String actRow2RateTestLocation			=batchPopUPRow2RateTxt.getText();
				String expRow2RateTestLocation			="1000";
				
				String actRow3BaseqtyTestLocation		=batchPopUPRow3BaseQuantityTxt.getText();
				String expRow3BaseqtyTestLocation		="3.000";
				
				String actRow3RateTestLocation			=batchPopUPRow3RateTxt.getText();
				String expRow3RateTestLocation			="2000";
				
				
			
				System.out.println("actRow1BatchNoTestLocation             :"+actRow1BatchNoTestLocation);
				System.out.println("expRow1BatchNoTestLocation             :"+expRow1BatchNoTestLocation);
				
				System.out.println("actRow2BatchNoTestLocation             :"+actRow2BatchNoTestLocation);
				System.out.println("expRow2BatchNoTestLocation             :"+expRow2BatchNoTestLocation);
				
				System.out.println("actRow3BatchNoTestLocation             :"+actRow3BatchNoTestLocation);
				System.out.println("expRow3BatchNoTestLocation             :"+expRow3BatchNoTestLocation);
				
				System.out.println("actRow3RateTestLocation                :"+actRow3RateTestLocation);
				System.out.println("expRow3RateTestLocation                :"+expRow3RateTestLocation);
				
				System.out.println("actRow3BaseqtyTestLocation             :"+actRow3BaseqtyTestLocation);
				System.out.println("expRow3BaseqtyTestLocation             :"+expRow3BaseqtyTestLocation);
				
				System.out.println("actRow1BaseqtyTestLocation             :"+actRow1BaseqtyTestLocation);
				System.out.println("expRow1BaseqtyTestLocation             :"+expRow1BaseqtyTestLocation);
				
				System.out.println("actRow1RateTestLocation                :"+actRow1RateTestLocation);
				System.out.println("expRow1RateTestLocation                :"+expRow1RateTestLocation);
				
				System.out.println("actRow2BaseqtyTestLocation             :"+actRow2BaseqtyTestLocation);
				System.out.println("expRow2BaseqtyTestLocation             :"+expRow2BaseqtyTestLocation);
				
				System.out.println("actRow2RateTestLocation                :"+actRow2RateTestLocation);
				System.out.println("expRow2RateTestLocation                :"+expRow2RateTestLocation);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
				
				boolean actBatchpopupIsDisplaying2 = batchPickOnFIFOIcon.isDisplayed();
				boolean expBatchpopupIsDisplaying2 = true;
				
				System.out.println("Batch Popup isDisplaying  : " + actBatchpopupIsDisplaying2 + "  Value Expected  "
						+ expBatchpopupIsDisplaying2);
				click(batchPickOnFIFOIcon);
				click(batchOkIcon);
				
				Thread.sleep(2000);
				click(select6thRow_1stColumn);

				int row3List=Row3ShortageInStockVoucherBodyList.size();
				System.err.println(row3List);
				ArrayList<String> row3ShortageStockList = new ArrayList<String>();
				for(int i=0 ; i < row3List ;i++)
			  	{
			  		String data=Row3ShortageInStockVoucherBodyList.get(i).getText();
			  		row3ShortageStockList.add(data);
			  	}
				
				String actrow3ShortageStockList = row3ShortageStockList.toString();
				String exprow3ShortageStockList = "[3, TEST, BR-BATCH, 2.000, 200.0000, 400.0000, 8.1111, 8.1111, 01082021PS, , ]";
				
				System.out.println("actrow3ShortageStockList       :"+actrow3ShortageStockList);		
				System.out.println("exprow3ShortageStockList       :"+exprow3ShortageStockList);
				
				
				int row4List=Row4ShortageInStockVoucherBodyList.size();
				System.err.println(row4List);
				ArrayList<String> row4ShortageStockList = new ArrayList<String>();
				for(int i=0 ; i < row4List ;i++)
			  	{
			  		String data=Row4ShortageInStockVoucherBodyList.get(i).getText();
			  		row4ShortageStockList.add(data);
			  	}
				
				String actrow4ShortageStockList = row4ShortageStockList.toString();
				String exprow4ShortageStockList = "[4, TEST, BR-BATCH, 3.000, 200.0000, 600.0000, 1000.0000, 1000.0000, 01092021PH, , ]";
				
				System.out.println("actrow4ShortageStockList       :"+actrow4ShortageStockList);		
				System.out.println("exprow4ShortageStockList       :"+exprow4ShortageStockList);
				
				int row5List=Row5ShortageInStockVoucherBodyList.size();
				System.err.println(row5List);
				ArrayList<String> row5ShortageStockList = new ArrayList<String>();
				for(int i=0 ; i < row5List ;i++)
			  	{
			  		String data=Row5ShortageInStockVoucherBodyList.get(i).getText();
			  		row5ShortageStockList.add(data);
			  	}
				
				String actrow5ShortageStockList = row5ShortageStockList.toString();
				String exprow5ShortageStockList = "[5, TEST, BR-BATCH, 1.000, 200.0000, 200.0000, 2000.0000, 2000.0000, 01092021PH, , ]";
				
				System.out.println("actrow5ShortageStockList       :"+actrow5ShortageStockList);		
				System.out.println("exprow5ShortageStockList       :"+exprow5ShortageStockList);
		
				if(actRow1BatchNoTestLocation.equalsIgnoreCase(expRow1BatchNoTestLocation) && actRow2BatchNoTestLocation.equalsIgnoreCase(expRow2BatchNoTestLocation) 
						&& actRow3BatchNoTestLocation.equalsIgnoreCase(expRow3BatchNoTestLocation) && actRow3BaseqtyTestLocation.equalsIgnoreCase(expRow3BaseqtyTestLocation) 
						&& actRow3RateTestLocation.equalsIgnoreCase(expRow3RateTestLocation) && actRow1BaseqtyTestLocation.equalsIgnoreCase(expRow1BaseqtyTestLocation)
						&& actRow1RateTestLocation.equalsIgnoreCase(expRow1RateTestLocation) && actRow2BaseqtyTestLocation.equalsIgnoreCase(expRow2BaseqtyTestLocation) 
						&& actRow2RateTestLocation.equalsIgnoreCase(expRow2RateTestLocation) && actrow3ShortageStockList.equalsIgnoreCase(exprow3ShortageStockList)
						&& actrow4ShortageStockList.equalsIgnoreCase(exprow4ShortageStockList) && actrow5ShortageStockList.equalsIgnoreCase(exprow5ShortageStockList)){
					
					return true;
				}
					else{
						return false;
					}
		
	          }
	
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[10]")
	private static WebElement  select6thRow_9thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[10]")
	private static WebElement  select7thRow_9thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[10]")
	private static WebElement  select8thRow_9thColumn;
	
	
		public boolean  checkSavingShortagesinStockVoucherWithFifoItem3WareHosue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{		
		
		
		// 3rd Row

				//click(select6thRow_1stColumn);
				Thread.sleep(2000);
				enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
				Thread.sleep(2000);
				enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
				enterText(enterWarehousettxt, "HYD");
				Thread.sleep(3000);
				tab(enterWarehousettxt);
				
				enterText(enterItemttxt, "FIFO-BIN");
				Thread.sleep(2000);
				tab(enterItemttxt);
				Thread.sleep(2000);	
				
				enterText(enterQuantitytxt, "3");
				tab(enterQuantitytxt);
				enterText(enterRatetxt, "100");
				tab(enterRatetxt);	
				
				click(select6thRow_9thColumn);
				
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

				boolean actBinpopupIsDisplaying = binSearchTxt.isDisplayed();
				boolean expBinpopupIsDisplaying = true;

				System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying + "  Value Expected  "
						+ expBinpopupIsDisplaying);

				click(binSearchBtn);
				Thread.sleep(2000);
				click(binPopup1stRow);
				Thread.sleep(1500);
				click(binPickBtn);
				//click(binAutoAllocateBtn);
				Thread.sleep(2000);
				click(binOkBtn);
				
				//7th Row
				
				click(select7thRow_1stColumn);
				Thread.sleep(2000);
				enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
				Thread.sleep(2000);
				enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
				enterText(enterWarehousettxt, "SEC");
				Thread.sleep(3000);
				tab(enterWarehousettxt);
				
				enterText(enterItemttxt, "FIFO-BIN");
				Thread.sleep(2000);
				tab(enterItemttxt);
				Thread.sleep(2000);	
				
				enterText(enterQuantitytxt, "3");
				tab(enterQuantitytxt);
				enterText(enterRatetxt, "100");
				tab(enterRatetxt);	
				
				click(select7thRow_9thColumn);
				
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

				boolean actBinpopupIsDisplaying1 = binSearchTxt.isDisplayed();
				boolean expBinpopupIsDisplaying1 = true;

				System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying1 + "  Value Expected  "
						+ expBinpopupIsDisplaying1);

				click(binSearchBtn);
				Thread.sleep(3000);
				click(binPopup1stRow);
				Thread.sleep(1500);
				click(binPickBtn);
				//click(binAutoAllocateBtn);
				Thread.sleep(2000);
				click(binOkBtn);
				
				Thread.sleep(2000);
				
				tab(bin1Txt);
				Thread.sleep(2000);
				
				//8th Row
				
				//click(select8thRow_1stColumn);
				Thread.sleep(2000);
				enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
				Thread.sleep(2000);
				enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
				enterText(enterWarehousettxt, "TEST");
				Thread.sleep(3000);
				tab(enterWarehousettxt);
				
				enterText(enterItemttxt, "FIFO-BIN");
				Thread.sleep(2000);
				tab(enterItemttxt);
				Thread.sleep(2000);	
				
				enterText(enterQuantitytxt, "6");
				tab(enterQuantitytxt);
				
				enterText(enterRatetxt, "200");
				tab(enterRatetxt);	
				
				click(select8thRow_9thColumn);
				
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

				boolean actBinpopupIsDisplaying2 = binSearchTxt.isDisplayed();
				boolean expBinpopupIsDisplaying2 = true;

				System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying2 + "  Value Expected  "
						+ expBinpopupIsDisplaying2);

				click(binSearchBtn);
				Thread.sleep(3000);
				click(binPopup1stRow);
				Thread.sleep(1500);
				click(binPickBtn);
				//click(binAutoAllocateBtn);
				Thread.sleep(2000);
				click(binOkBtn);
				Thread.sleep(2000);
				tab(bin1Txt);
				Thread.sleep(2000);

				if(actBinpopupIsDisplaying==expBinpopupIsDisplaying && actBinpopupIsDisplaying1==expBinpopupIsDisplaying1 && actBinpopupIsDisplaying2==expBinpopupIsDisplaying2){
					return true;
					
				}else{
					return false;
				}
				
		}
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[11]")
		private static WebElement  select9thRow_10thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[11]")
		private static WebElement  select10thRow_10thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[11]")
		private static WebElement  select11thRow_10thColumn;
		
		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td")
		private static List<WebElement> Row9ShortageInStockVoucherBodyList;
		
		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td")
		private static List<WebElement> Row10ShortageInStockVoucherBodyList;
		
		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td")
		private static List<WebElement> Row11ShortageInStockVoucherBodyList;
		
		
		public boolean  checkSavingShortagesinStockVoucherWithRmaItemHydWareHosue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{		
			
			
		// 9th Row

		//click(select9thRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(3000);
		tab(enterWarehousettxt);
		
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);
		
		enterText(enterQuantitytxt, "3");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		
		click(select9thRow_10thColumn);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying + "  Value Expected  "
				+ expRMAPopupIsDisplaying);
		
		click(rmaSearchGridField);
		Thread.sleep(1000);
		
		int count = rmaSerialNos.size();
		
		ArrayList<String> SearchRMA = new ArrayList<String>();
		
		for (int i = 0; i < count; i++) 
		{
			String data = rmaSerialNos.get(i).getText();
			SearchRMA.add(data);
			
			if (data.equalsIgnoreCase("01092021PS7") || data.equalsIgnoreCase("01092021PS8") || data.equalsIgnoreCase("01092021PS9")) 
			{
				rmaSerialNoCheckBoxes.get(i).click();
			}
		}
		
		String actSearchRMASerialNoList = SearchRMA.toString();
		String expSearchRMASerialNoList = "[01092021PS7, 01092021PS8, 01092021PS9, 04092021EXH1, 25022021EXH1]"; 
		
		System.out.println("actSearchRMASerialNoList       :"+actSearchRMASerialNoList);
		System.out.println("expSearchRMASerialNoList       :"+expSearchRMASerialNoList);
		
		
		//click(searchRmaSelectAllGridCheckBox);
		click(searchRmaOkBtn);
		
		int serialcount	=rmaSerialNoList.size();
		ArrayList<String> rmaSerialnos = new ArrayList<String>();
		for (int i = 0; i < serialcount; i++) {
			String data=rmaSerialNoList.get(i).getText();
			rmaSerialnos.add(data);
		}
		
		String actRmaSerialNo 	=rmaSerialnos.toString();
		String expRmaSerialNo	="[01092021PS7, 01092021PS8, 01092021PS9]";
		
		System.out.println("actRmaSerialNo       :"+actRmaSerialNo);
		System.out.println("expRmaSerialNo       :"+expRmaSerialNo);
		
	
		click(RMAPopupOkBtn);	
		Thread.sleep(2000);
		tab(enterRmatxt);
		
		Thread.sleep(2000);
		int row9List=Row9ShortageInStockVoucherBodyList.size();
		System.err.println(row9List);
		ArrayList<String> row9ShortageStockList = new ArrayList<String>();
		for(int i=0 ; i < row9List ;i++)
	  	{
	  		String data=Row9ShortageInStockVoucherBodyList.get(i).getText();
	  		row9ShortageStockList.add(data);
	  	}
		
		String actrow9ShortageStockList = row9ShortageStockList.toString();
		String exprow9ShortageStockList = "[9, HYD, WA-RMA, 3.000, 100.0000, 300.0000, 687.9324, 1135.2497, , , 01092021PS7,01092021PS8,01092021PS9]";
		
		System.out.println("actrow9ShortageStockList       :"+actrow9ShortageStockList);		
		System.out.println("exprow9ShortageStockList       :"+exprow9ShortageStockList);
		
		
		
		
		
		if(actRMAPopupIsDisplaying==expRMAPopupIsDisplaying && actSearchRMASerialNoList.equalsIgnoreCase(expSearchRMASerialNoList) 
				&& actRmaSerialNo.equalsIgnoreCase(expRmaSerialNo) && actrow9ShortageStockList.equalsIgnoreCase(exprow9ShortageStockList)){
			return true;
			
		}else{
			return false;
		}


				
		}
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[2]")
		private static WebElement  select10thRow_1stColumn;
		
		public boolean  checkSavingShortagesinStockVoucherWithRmaItemSecWareHosue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{		
			
			
			// 10th Row

			click(select10thRow_1stColumn);
			Thread.sleep(2000);
			enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(2000);
			enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
			enterText(enterWarehousettxt, "SEC");
			Thread.sleep(3000);
			tab(enterWarehousettxt);
			
			enterText(enterItemttxt, "WA-RMA");
			Thread.sleep(2000);
			tab(enterItemttxt);
			Thread.sleep(2000);
			
			enterText(enterQuantitytxt, "3");
			tab(enterQuantitytxt);
			enterText(enterRatetxt, "100");
			tab(enterRatetxt);
			
			click(select10thRow_10thColumn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

			boolean actRMAPopupIsDisplaying = rmaScreenTitle.isDisplayed();
			boolean expRMAPopupIsDisplaying = true;

			System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying + "  Value Expected  "
					+ expRMAPopupIsDisplaying);
			
			click(rmaSearchGridField);
			Thread.sleep(1000);
			
			int count = rmaSerialNos.size();
			
			ArrayList<String> SearchRMA = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data = rmaSerialNos.get(i).getText();
				SearchRMA.add(data);
				
				if (data.equalsIgnoreCase("01092021PH17") || data.equalsIgnoreCase("01092021PH18") || data.equalsIgnoreCase("04082021EXH1")) 
				{
					rmaSerialNoCheckBoxes.get(i).click();
				}
			}
			
			String actSearchRMASerialNoList = SearchRMA.toString();
			String expSearchRMASerialNoList = "[01092021PH17, 01092021PH18, 04082021EXH1, 04092021EXS1]"; 
			
			System.out.println("actSearchRMASerialNoList       :"+actSearchRMASerialNoList);
			System.out.println("expSearchRMASerialNoList       :"+expSearchRMASerialNoList);
			
			
			//click(searchRmaSelectAllGridCheckBox);
			click(searchRmaOkBtn);
			
			int serialcount	=rmaSerialNoList.size();
			ArrayList<String> rmaSerialnos = new ArrayList<String>();
			for (int i = 0; i < serialcount; i++) {
				String data=rmaSerialNoList.get(i).getText();
				rmaSerialnos.add(data);
			}
			
			String actRmaSerialNo 	=rmaSerialnos.toString();
			String expRmaSerialNo	="[01092021PH17, 01092021PH18, 04082021EXH1]";
			
			System.out.println("actRmaSerialNo       :"+actRmaSerialNo);
			System.out.println("expRmaSerialNo       :"+expRmaSerialNo);
			
		
			click(RMAPopupOkBtn);	
			Thread.sleep(2000);
			tab(enterRmatxt);
			
			Thread.sleep(2000);
			int row10List=Row10ShortageInStockVoucherBodyList.size();
			System.err.println(row10List);
			ArrayList<String> row10ShortageStockList = new ArrayList<String>();
			for(int i=0 ; i < row10List ;i++)
		  	{
		  		String data=Row10ShortageInStockVoucherBodyList.get(i).getText();
		  		row10ShortageStockList.add(data);
		  	}
			
			String actrow10ShortageStockList = row10ShortageStockList.toString();
			String exprow10ShortageStockList = "[10, SEC, WA-RMA, 3.000, 100.0000, 300.0000, 1414.2531, 1135.2497, , , 01092021PH17,01092021PH18,04082021EXH1]";
			
			System.out.println("actrow10ShortageStockList       :"+actrow10ShortageStockList);		
			System.out.println("exprow10ShortageStockList       :"+exprow10ShortageStockList);
			
			
			if(actRMAPopupIsDisplaying==expRMAPopupIsDisplaying && actSearchRMASerialNoList.equalsIgnoreCase(expSearchRMASerialNoList) 
					&& actRmaSerialNo.equalsIgnoreCase(expRmaSerialNo) && actrow10ShortageStockList.equalsIgnoreCase(exprow10ShortageStockList)){
				return true;
				
			}else{
				return false;
			}


					
			}
		
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[2]")
		private static WebElement  select11thRow_1stColumn;
		
		public boolean  checkSavingShortagesinStockVoucherWithRmaItemTestWareHosue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{		
			
			
			// 11th Row

			click(select11thRow_1stColumn);
			Thread.sleep(2000);
			enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(2000);
			enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
			enterText(enterWarehousettxt, "TEST");
			Thread.sleep(3000);
			tab(enterWarehousettxt);
			
			enterText(enterItemttxt, "WA-RMA");
			Thread.sleep(2000);
			tab(enterItemttxt);
			Thread.sleep(2000);
			
			enterText(enterQuantitytxt, "6");
			tab(enterQuantitytxt);
			enterText(enterRatetxt, "200");
			tab(enterRatetxt);

			click(select11thRow_10thColumn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

			boolean actRMAPopupIsDisplaying = rmaScreenTitle.isDisplayed();
			boolean expRMAPopupIsDisplaying = true;

			System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying + "  Value Expected  "
					+ expRMAPopupIsDisplaying);
			
			click(rmaSearchGridField);
			Thread.sleep(1000);
			
			int count = rmaSerialNos.size();
			
			ArrayList<String> SearchRMA = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data = rmaSerialNos.get(i).getText();
				SearchRMA.add(data);
				
				if (data.equalsIgnoreCase("01092021PS4") || data.equalsIgnoreCase("01092021PS5") || data.equalsIgnoreCase("01092021PS6")
					|| data.equalsIgnoreCase("01092021PH14") || data.equalsIgnoreCase("01092021PH15") || data.equalsIgnoreCase("01092021PH16")	) 
				{
					rmaSerialNoCheckBoxes.get(i).click();
				}
			}
			
			String actSearchRMASerialNoList = SearchRMA.toString();
			String expSearchRMASerialNoList = "[01092021PS4, 01092021PS5, 01092021PS6, 01092021PH14, 01092021PH15, 01092021PH16, 01102020PT10, 06102020PT10]"; 
			
			System.out.println("actSearchRMASerialNoList       :"+actSearchRMASerialNoList);
			System.out.println("expSearchRMASerialNoList       :"+expSearchRMASerialNoList);
			
			
			//click(searchRmaSelectAllGridCheckBox);
			click(searchRmaOkBtn);
			
			int serialcount	=rmaSerialNoList.size();
			ArrayList<String> rmaSerialnos = new ArrayList<String>();
			for (int i = 0; i < serialcount; i++) {
				String data=rmaSerialNoList.get(i).getText();
				rmaSerialnos.add(data);
			}
			
			String actRmaSerialNo 	=rmaSerialnos.toString();
			String expRmaSerialNo	="[01092021PS4, 01092021PS5, 01092021PS6, 01092021PH14, 01092021PH15, 01092021PH16]";
			
			System.out.println("actRmaSerialNo       :"+actRmaSerialNo);
			System.out.println("expRmaSerialNo       :"+expRmaSerialNo);
			
		
			click(RMAPopupOkBtn);	
			Thread.sleep(2000);
			tab(enterRmatxt);
			
			Thread.sleep(2000);
			int row11List=Row11ShortageInStockVoucherBodyList.size();
			System.err.println(row11List);
			ArrayList<String> row11ShortageStockList = new ArrayList<String>();
			for(int i=0 ; i < row11List ;i++)
		  	{
		  		String data=Row11ShortageInStockVoucherBodyList.get(i).getText();
		  		row11ShortageStockList.add(data);
		  	}
			
			String actrow11ShortageStockList = row11ShortageStockList.toString();
			String exprow11ShortageStockList = "[11, TEST, WA-RMA, 6.000, 200.0000, 1,200.0000, 998.0464, 1135.2497, , , 01092021PS4,01092021PS5,01092021PS6,01092021PH14,01092021PH15,01092021PH16]";
			
			System.out.println("actrow11ShortageStockList       :"+actrow11ShortageStockList);		
			System.out.println("exprow11ShortageStockList       :"+exprow11ShortageStockList);
			
			
			Thread.sleep(2000);
			click(voucherSaveBtn);

			HashSet<String> actMessage = new HashSet<String>();

			for (int i = 0; i < 2; i++) {
				String data = checkValidationMessage("");
				actMessage.add(data);
			}

			HashSet<String> expMessage = new HashSet<String>();

			expMessage.add("Voucher saved successfully : 33");
			expMessage.add("Saving in background.");

			System.out.println("Actual Message    : " + actMessage);
			System.out.println("Expected Message  : " + expMessage);
			
			
			if(actRMAPopupIsDisplaying==expRMAPopupIsDisplaying && actSearchRMASerialNoList.equalsIgnoreCase(expSearchRMASerialNoList) 
					&& actRmaSerialNo.equalsIgnoreCase(expRmaSerialNo) && actMessage.equals(expMessage) && actrow11ShortageStockList.equalsIgnoreCase(exprow11ShortageStockList)){
				click(new_CloseBtn);			
				Thread.sleep(1000);
				click(voucherhomeCloseBtn);
				return true;
				
			}else{
				
				click(new_CloseBtn);			
				Thread.sleep(1000);
				click(voucherhomeCloseBtn);
				return false;
			}


					
			}
		
		@FindBy(xpath="//*[@id='2009']/span")
		private static WebElement  financialsTransactionsPurchaseMenuPurchasesReturns;
		
		@FindBy(xpath="//input[@id='id_header_2']")
		private static WebElement  purchaseReturnsDatetxt;
		
		@FindBy(xpath="//input[@id='id_header_3']")
		private static WebElement  purchasesReturnsAccounttxt;
		
		@FindBy(xpath="//input[@id='id_header_4']")
		private static WebElement  purchasesReturnsvendorAccounttxt;
		
		@FindBy(xpath="//input[@id='id_header_268435459']")
		private static WebElement  purchasesReturnsDepartmenttxt;
		
		public boolean  checkSavingPurchaseReturnsVoucherWithBrItemHydSecWareHosueInINVUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{
			
			Thread.sleep(3000);
			click(userNameDisplay);
			Thread.sleep(1000);
			click(logoutOption);		
			Thread.sleep(3000);
			
			String unamelt="inv";

			String pawslt="su";
			
			LoginPage.enterUserName(unamelt);	     
	        LoginPage.enterPassword(pawslt);
	        
			
			 Select company = new Select(companyDropDownList);
			 company.selectByValue("108");
			 Thread.sleep(2000);
			 
	  
	        LoginPage.clickOnSignInBtn();  
	        
	        Thread.sleep(3000);
	        
	        System.out.println("Signing");
	        
	        // Login User Name is Verify
	        
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo                      = userNameDisplay.getText();
			String expUserInfo                      = "INV";	

			System.out.println("Login With Valid User and Check User Name Display Value Actual : ." + actUserInfo + ". Value Expected : " + expUserInfo);
			
			
			// Login Company Name is Verify
			Thread.sleep(1000);
			click(companyLogo);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
			String actGetLoginCompanyNameInformation = companyName.getText();
			String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
			String expGetLoginCompanyName            = "Maintain Quantity a";   
			companyLogo.click();
			
			System.out.println("Company Name Display Value Actual                : ." + actGetLoginCompanyName +  ". Value Expected : " + expGetLoginCompanyName);
			
			
			Thread.sleep(1000);
			
			click(financialsMenu);
			click(financialsTransactionMenu);
			click(financialsTransactionsPurchaseMenu);
			click(financialsTransactionsPurchaseMenuPurchasesReturns);
			
			Thread.sleep(2000);
			
			click(newBtn);	
			
			checkValidationMessage("Screen opened");
			
			click(purchaseReturnsDatetxt);
			purchaseReturnsDatetxt.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(2000);
			purchaseReturnsDatetxt.sendKeys(Keys.BACK_SPACE);
			enterText(purchaseReturnsDatetxt, "06/09/2021");
			tab(purchaseReturnsDatetxt);
			
			Thread.sleep(2000);
			enterText(purchasesReturnsAccounttxt, "Cost of goods sold - Computers");
			tab(purchasesReturnsAccounttxt);
			Thread.sleep(2000);
			enterText(purchasesReturnsvendorAccounttxt, "Bank");
			Thread.sleep(2000);
			click(purchasesReturnsDepartmenttxt);
			enterText(purchasesReturnsDepartmenttxt, "DEPT1");
			tab(purchasesReturnsDepartmenttxt);			
			Thread.sleep(2000);
			
			click(select1stRow_1stColumn);
			enterText(enterWarehousettxt, "HYD");
			Thread.sleep(2000);
			tab(enterWarehousettxt);
			
			enterText(enterItemttxt, "BR-BATCH");
			Thread.sleep(2000);
			tab(enterItemttxt);
			
			enterText(enterQuantitytxt, "1");
			tab(enterQuantitytxt);
			enterText(enterRatetxt, "100");
			tab(enterRatetxt);
			Thread.sleep(2000);
			
			click(select1stRow_9thColumn);			
			Thread.sleep(2000);
			
			String actRow1BatchNo		=batchPopUPRow1BatchNoTxt.getText();
			String expRow1BatchNo		="01092021PH";
			
			String actRow2BatchNo		=batchPopUPRow2BatchNoTxt.getText();
			String expRow2BatchNo		="04092021EXH";
			
			String actRow1Baseqty		=batchPopUPRow1BaseQuantityTxt.getText();
			String expRow1Baseqty		="1.000";
			
			String actRow1Rate			=batchPopUPRow1RateTxt.getText();
			String expRow1Rate			="1000";
			
			String actRow2Baseqty		=batchPopUPRow2BaseQuantityTxt.getText();
			String expRow2Baseqty		="1.000";
			
			String actRow2Rate			=batchPopUPRow2RateTxt.getText();
			String expRow2Rate			="100";
			
		
			System.out.println("actRow1BatchNo             :"+actRow1BatchNo);
			System.out.println("expRow1BatchNo             :"+expRow1BatchNo);
			
			System.out.println("actRow2BatchNo             :"+actRow2BatchNo);
			System.out.println("expRow2BatchNo             :"+expRow2BatchNo);		
			
			
			System.out.println("actRow1Baseqty             :"+actRow1Baseqty);
			System.out.println("expRow1Baseqty             :"+expRow1Baseqty);
			
			System.out.println("actRow1Rate                :"+actRow1Rate);
			System.out.println("expRow1Rate                :"+expRow1Rate);
			
			System.out.println("actRow2Baseqty             :"+actRow2Baseqty);
			System.out.println("expRow2Baseqty             :"+expRow2Baseqty);
			
			System.out.println("actRow2Rate                :"+actRow2Rate);
			System.out.println("expRow2Rate                :"+expRow2Rate);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
			
			boolean actBatchpopupIsDisplaying = batchPickOnFIFOIcon.isDisplayed();
			boolean expBatchpopupIsDisplaying = true;
			
			System.out.println("Batch Popup isDisplaying  : " + actBatchpopupIsDisplaying + "  Value Expected  "
					+ expBatchpopupIsDisplaying);
			click(batchPickOnFIFOIcon);
			click(batchOkIcon);

			// 2nd Row
			
			click(select2ndRow_1stColumn);
			Thread.sleep(2000);
			enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(2000);
			enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
			enterText(enterWarehousettxt, "SEC");
			Thread.sleep(3000);
			tab(enterWarehousettxt);
			
			enterText(enterItemttxt, "BR-BATCH");
			Thread.sleep(2000);
			tab(enterItemttxt);
			Thread.sleep(2000);	
			
			enterText(enterQuantitytxt, "1");
			tab(enterQuantitytxt);
			enterText(enterRatetxt, "100");
			tab(enterRatetxt);
			
			click(select2ndRow_9thColumn);
			
			Thread.sleep(2000);
			
			String actRow1BatchNoSecLocation		=batchPopUPRow1BatchNoTxt.getText();
			String expRow1BatchNoSecLocation		="04092021EXS";
		
			
			String actRow1BaseqtySecLocation		=batchPopUPRow1BaseQuantityTxt.getText();
			String expRow1BaseqtySecLocation		="1.000";
			
			String actRow1RateSecLocation			=batchPopUPRow1RateTxt.getText();
			String expRow1RateSecLocation			="200";
			
		
		
			System.out.println("actRow1BatchNoSecLocation             :"+actRow1BatchNoSecLocation);
			System.out.println("expRow1BatchNoSecLocation             :"+expRow1BatchNoSecLocation);			
			
			System.out.println("actRow1BaseqtySecLocation             :"+actRow1BaseqtySecLocation);
			System.out.println("expRow1BaseqtySecLocation             :"+expRow1BaseqtySecLocation);
			
			System.out.println("actRow1RateSecLocation                :"+actRow1RateSecLocation);
			System.out.println("expRow1RateSecLocation                :"+expRow1RateSecLocation);
			
		
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
			
			boolean actBatchpopupIsDisplaying1 = batchPickOnFIFOIcon.isDisplayed();
			boolean expBatchpopupIsDisplaying1 = true;
			
			System.out.println("Batch Popup isDisplaying  : " + actBatchpopupIsDisplaying1 + "  Value Expected  "
					+ expBatchpopupIsDisplaying1);
			click(batchPickOnFIFOIcon);
			click(batchOkIcon);
			
			if(actRow1BatchNo.equalsIgnoreCase(expRow1BatchNo) && actRow2BatchNo.equalsIgnoreCase(expRow2BatchNo) 
					&& actRow1Baseqty.equalsIgnoreCase(expRow1Baseqty) && actRow1Rate.equalsIgnoreCase(expRow1Rate) 
					&& actRow2Baseqty.equalsIgnoreCase(expRow2Baseqty) && actRow2Rate.equalsIgnoreCase(expRow2Rate)
					&& actRow1BatchNoSecLocation.equalsIgnoreCase(expRow1BatchNoSecLocation) 
					&& actRow1BaseqtySecLocation.equalsIgnoreCase(expRow1BaseqtySecLocation) 
					&& actRow1RateSecLocation.equalsIgnoreCase(expRow1RateSecLocation)){
				
				return true;
			}
				else{
					return false;
				} 
		
		
		}
		
		
		public boolean  checkSavingPurchaseReturnsVoucherWithFifoItemHydSecWareHosueInINVUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{		
			
			
		// 3rd Row

		click(select3rdRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(3000);
		tab(enterWarehousettxt);
		
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);
		
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		
		click(select3rdRow_10thColumn);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying + "  Value Expected  "
				+ expBinpopupIsDisplaying);

		click(binSearchBtn);
		Thread.sleep(3000);
		click(binPopup1stRow);
		Thread.sleep(1500);
		click(binPickBtn);
		//click(binAutoAllocateBtn);
		Thread.sleep(2000);
		click(binOkBtn);

		// 4th Row

		click(select4thRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(3000);
		tab(enterWarehousettxt);
		
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);
		
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		
		
		click(select4thRow_10thColumn);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying1 = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying1 = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying1 + "  Value Expected  "
				+ expBinpopupIsDisplaying1);

		click(binSearchBtn);
		Thread.sleep(3000);
		click(binPopup1stRow);
		Thread.sleep(1500);
		click(binPickBtn);
		//click(binAutoAllocateBtn);
		Thread.sleep(2000);
		click(binOkBtn);

		if (actBinpopupIsDisplaying == expBinpopupIsDisplaying
				&& actBinpopupIsDisplaying1 == expBinpopupIsDisplaying1) {
			return true;

		} else {
			return false;
		}

	}
		
		public boolean  checkSavingPurchaseReturnsVoucherWithRmaItemHydWareHosueInINVUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{		
			
			
		// 5th Row

		click(select5thRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(3000);
		tab(enterWarehousettxt);
		
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);
		
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		
		click(select5thRow_11thColumn);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying + "  Value Expected  "
				+ expRMAPopupIsDisplaying);
		
		click(rmaSearchGridField);
		Thread.sleep(1000);
		
		int count = rmaSerialNos.size();
		
		ArrayList<String> SearchRMA = new ArrayList<String>();
		
		for (int i = 0; i < count; i++) 
		{
			String data = rmaSerialNos.get(i).getText();
			SearchRMA.add(data);
			
			if (data.equalsIgnoreCase("04092021EXH1")) 
			{
				rmaSerialNoCheckBoxes.get(i).click();
			}
		}
		
		String actSearchRMASerialNoList = SearchRMA.toString();
		String expSearchRMASerialNoList = "[04092021EXH1, 25022021EXH1]"; 
		
		System.out.println("actSearchRMASerialNoList       :"+actSearchRMASerialNoList);
		System.out.println("expSearchRMASerialNoList       :"+expSearchRMASerialNoList);
		
		
		//click(searchRmaSelectAllGridCheckBox);
		click(searchRmaOkBtn);
		
		int serialcount	=rmaSerialNoList.size();
		ArrayList<String> rmaSerialnos = new ArrayList<String>();
		for (int i = 0; i < serialcount; i++) {
			String data=rmaSerialNoList.get(i).getText();
			rmaSerialnos.add(data);
		}
		
		String actRmaSerialNo 	=rmaSerialnos.toString();
		String expRmaSerialNo	="[04092021EXH1]";
		
		System.out.println("actRmaSerialNo       :"+actRmaSerialNo);
		System.out.println("expRmaSerialNo       :"+expRmaSerialNo);
		
	
		click(RMAPopupOkBtn);	
		Thread.sleep(2000);
		tab(enterRmatxt);
		
		Thread.sleep(2000);
		int row5List=Row5ShortageInStockVoucherBodyList.size();
		System.err.println(row5List);
		ArrayList<String> row5ShortageStockList = new ArrayList<String>();
		for(int i=0 ; i < row5List ;i++)
	  	{
	  		String data=Row5ShortageInStockVoucherBodyList.get(i).getText();
	  		row5ShortageStockList.add(data);
	  	}
		
		String actrow5ShortageStockList = row5ShortageStockList.toString();
		String exprow5ShortageStockList = "[5, HYD, WA-RMA, 1.000, 687.9324, 687.9324, 0.00, 1048.7673000, 687.9324000, , , 04092021EXH1]";
		
		System.out.println("actrow5ShortageStockList       :"+actrow5ShortageStockList);		
		System.out.println("exprow5ShortageStockList       :"+exprow5ShortageStockList);
		
		
		
		
		
		if(actRMAPopupIsDisplaying==expRMAPopupIsDisplaying && actSearchRMASerialNoList.equalsIgnoreCase(expSearchRMASerialNoList) 
				&& actRmaSerialNo.equalsIgnoreCase(expRmaSerialNo) && actrow5ShortageStockList.equalsIgnoreCase(exprow5ShortageStockList)){
			return true;
			
		}else{
			return false;
		}


				
		}
		
		public boolean  checkSavingPurchaseReturnsVoucherWithRmaItemSECWareHosueInINVUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException{		
			
			
		// 6th Row

		click(select6thRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(3000);
		tab(enterWarehousettxt);
		
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);
		
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "100");
		tab(enterRatetxt);
		
		click(select6thRow_11thColumn);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying + "  Value Expected  "
				+ expRMAPopupIsDisplaying);
		
		click(rmaSearchGridField);
		Thread.sleep(1000);
		
		int count = rmaSerialNos.size();
		
		ArrayList<String> SearchRMA = new ArrayList<String>();
		
		for (int i = 0; i < count; i++) 
		{
			String data = rmaSerialNos.get(i).getText();
			SearchRMA.add(data);
			
			if (data.equalsIgnoreCase("04092021EXS1")) 
			{
				rmaSerialNoCheckBoxes.get(i).click();
			}
		}
		
		String actSearchRMASerialNoList = SearchRMA.toString();
		String expSearchRMASerialNoList = "[04092021EXS1]"; 
		
		System.out.println("actSearchRMASerialNoList       :"+actSearchRMASerialNoList);
		System.out.println("expSearchRMASerialNoList       :"+expSearchRMASerialNoList);
		
		
		//click(searchRmaSelectAllGridCheckBox);
		click(searchRmaOkBtn);
		
		int serialcount	=rmaSerialNoList.size();
		ArrayList<String> rmaSerialnos = new ArrayList<String>();
		for (int i = 0; i < serialcount; i++) {
			String data=rmaSerialNoList.get(i).getText();
			rmaSerialnos.add(data);
		}
		
		String actRmaSerialNo 	=rmaSerialnos.toString();
		String expRmaSerialNo	="[04092021EXS1]";
		
		System.out.println("actRmaSerialNo       :"+actRmaSerialNo);
		System.out.println("expRmaSerialNo       :"+expRmaSerialNo);
		
	
		click(RMAPopupOkBtn);	
		Thread.sleep(2000);
		tab(enterRmatxt);
		
		Thread.sleep(2000);
		int row6List=Row6ShortageInStockVoucherBodyList.size();
		System.err.println(row6List);
		ArrayList<String> row6ShortageStockList = new ArrayList<String>();
		for(int i=0 ; i < row6List ;i++)
	  	{
	  		String data=Row6ShortageInStockVoucherBodyList.get(i).getText();
	  		row6ShortageStockList.add(data);
	  	}
		
		String actrow6ShortageStockList = row6ShortageStockList.toString();
		String exprow6ShortageStockList = "[6, SEC, WA-RMA, 1.000, 1,414.2531, 1,414.2531, 0.00, 1048.7673000, 1414.2531000, , , 04092021EXS1]";
		
		System.out.println("actrow6ShortageStockList       :"+actrow6ShortageStockList);		
		System.out.println("exprow6ShortageStockList       :"+exprow6ShortageStockList);
		
		Thread.sleep(2000);
		click(voucherSaveBtn);

		HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) {
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : 33");
		expMessage.add("Saving in background.");

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage);
		
		
		
		if(actRMAPopupIsDisplaying==expRMAPopupIsDisplaying && actSearchRMASerialNoList.equalsIgnoreCase(expSearchRMASerialNoList) 
				&& actRmaSerialNo.equalsIgnoreCase(expRmaSerialNo) && actrow6ShortageStockList.equalsIgnoreCase(exprow6ShortageStockList)
				&& actMessage.equals(expMessage) ){
			click(new_CloseBtn);			
			Thread.sleep(1000);
			click(voucherhomeCloseBtn);
			
			return true;
			
		}else{
			
			click(new_CloseBtn);			
			Thread.sleep(1000);
			click(voucherhomeCloseBtn);
			
			return false;
		}
				
	}
		
		
		@FindBy(xpath="//*[@id='2016']/span")
		private static WebElement  financialsTransactionsSalesMenuSalesReturn; 
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[9]")
		private static WebElement  select5thRow_SalesReturnsAvgRate;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[8]")
		private static WebElement  select5thRow_SalesReturnsOverallAvgRate;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[9]")
		private static WebElement  select6thRow_SalesReturnsAvgRate;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[8]")
		private static WebElement  select6thRow_SalesReturnsOverallAvgRate;
		
		public boolean  checkSavingSalesReturnsVoucherWithAllItemsINVUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			
			
	    click(financialsMenu);
		click(financialsTransactionMenu);
		click(financialsTransactionsSalesMenu);
		click(financialsTransactionsSalesMenuSalesReturn);
		Thread.sleep(2000);
		
		click(newBtn);

		checkValidationMessage("Screen opened");

		click(salesDatetxt);
		salesDatetxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		salesDatetxt.sendKeys(Keys.BACK_SPACE);
		enterText(salesDatetxt, "07/09/2021");
		Thread.sleep(2000);
		enterText(salesAccounttxt, "Sales - Computers");
		tab(salesAccounttxt);
		Thread.sleep(2000);
		enterText(customerAccounttxt, "Bank");
		Thread.sleep(2000);
		click(customerAccounttxt);
		enterText(salesDepartmenttxt, "DEPT1");
		tab(salesDepartmenttxt);
		Thread.sleep(3000);

		//1st Row
		
		click(select1stRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(3000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "BR-BATCH");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);

		
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "1");
		tab(enterRatetxt);
		
		click(select1stRow_9thColumn);
		
		enterText(enterBatchtxt, "01102020PH");
		tab(enterBatchtxt);

		// 2nd Row

		click(select2ndRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(3000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "BR-BATCH");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);

		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "1");
		tab(enterRatetxt);

		click(select2ndRow_9thColumn);

		enterText(enterBatchtxt, "01102020PS");
		tab(enterBatchtxt);

		// 3rd Row

		click(select3rdRow_1stColumn);
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(3000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		Thread.sleep(2000);

		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "1");
		tab(enterRatetxt);

		click(select3rdRow_10thColumn);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying + "  Value Expected  "
				+ expBinpopupIsDisplaying);

		click(binSearchBtn);
		Thread.sleep(3000);
		click(binPopup1stRow);
		Thread.sleep(2000);
		click(binPickBtn);
		Thread.sleep(1000);
		click(binOkBtn);
		
		
		// 4th Row
		
		click(select4thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "FIFO-BIN");
		Thread.sleep(2000);
		tab(enterItemttxt);
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "1");
		tab(enterRatetxt);
		click(select4thRow_10thColumn);
		
	
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		boolean actBinpopupIsDisplaying1 = binSearchTxt.isDisplayed();
		boolean expBinpopupIsDisplaying1 = true;

		System.out.println("Bin Popup isDisplaying  : " + actBinpopupIsDisplaying1 + "  Value Expected  "
				+ expBinpopupIsDisplaying1);

		click(binSearchBtn);
		Thread.sleep(3000);
		click(binPopup1stRow);
		Thread.sleep(2000);
		click(binPickBtn);
		Thread.sleep(1000);
		click(binOkBtn);
		Thread.sleep(2000);
		

		// 5th Row

		click(select5thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "HYD");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);
		tab(enterItemttxt);
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "1");
		tab(enterRatetxt);
		click(select5thRow_11thColumn);
		
		String actual5throwAvgRate 			= select5thRow_SalesReturnsAvgRate.getText();
		String exp5throwAvgRate			    = "687.9324000";
		String actual5throwOverallAvgRate   = select5thRow_SalesReturnsOverallAvgRate.getText();
		String exp5throwoverallAvgRate      = "1048.7691000";

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying + "  Value Expected  "
				+ expRMAPopupIsDisplaying);
		
		enterText(rmaSerialNumberTxtField, "01102020PH1,1");
		click(rmaAddBtn);
		Thread.sleep(2000);
		click(RMAPopupOkBtn);
		Thread.sleep(3000);
		tab(enterRmatxt);
		//tab(enterBatCodetxt);

		// 6th Row
		Thread.sleep(1000);
		click(select6thRow_1stColumn);
		enterWarehousettxt.sendKeys(Keys.CONTROL, "a");
		Thread.sleep(2000);
		enterWarehousettxt.sendKeys(Keys.BACK_SPACE);
		enterText(enterWarehousettxt, "SEC");
		Thread.sleep(2000);
		tab(enterWarehousettxt);
		enterText(enterItemttxt, "WA-RMA");
		Thread.sleep(2000);
		tab(enterItemttxt);
		enterText(enterQuantitytxt, "1");
		tab(enterQuantitytxt);
		enterText(enterRatetxt, "1");
		tab(enterRatetxt);
		click(select6thRow_11thColumn);
		
		String actual6throwAvgRate 				= select6thRow_SalesReturnsAvgRate.getText();
		String exp6throwAvgRate 				= "1414.2531000";
		String actual6throwOverallAvgRate		= select6thRow_SalesReturnsOverallAvgRate.getText();
		String exp6throwoverallAvgRate			= "1048.7691000";

		System.out.println("actual5throwAvgRate       	 :" + actual5throwAvgRate);
		System.out.println("exp5throwAvgRate             :" + exp5throwAvgRate);

		System.out.println("actual5throwOverallAvgRate   :" + actual5throwOverallAvgRate);
		System.out.println("exp5throwoverallAvgRate      	:" + exp5throwoverallAvgRate);

		System.out.println("actual6throwAvgRate      	 :" + actual6throwAvgRate);
		System.out.println("exp6throwAvgRate          	 :" + exp6throwAvgRate);

		System.out.println("actual6throwOverallAvgRate   :" + actual6throwOverallAvgRate);
		System.out.println("exp6throwoverallAvgRate      :" + exp6throwoverallAvgRate);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		boolean actRMAPopupIsDisplaying1 = rmaScreenTitle.isDisplayed();
		boolean expRMAPopupIsDisplaying1 = true;

		System.out.println("RMA Popup isDisplaying  : " + actRMAPopupIsDisplaying1 + "  Value Expected  "
				+ expRMAPopupIsDisplaying1);
		enterText(rmaSerialNumberTxtField, "01102020PS1,1");
		click(rmaAddBtn);
		Thread.sleep(2000);
		click(RMAPopupOkBtn);
		Thread.sleep(2000);		
		tab(enterRmatxt);
		
		
		Thread.sleep(3000);
		click(voucherSaveBtn);

		HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) {
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : 34");
		expMessage.add("Saving in background.");

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage);

		if (actMessage.equals(expMessage) && actual5throwAvgRate.equalsIgnoreCase(exp5throwAvgRate)
				&& actual5throwOverallAvgRate.equalsIgnoreCase(exp5throwoverallAvgRate)
				&& actual6throwAvgRate.equalsIgnoreCase(exp6throwAvgRate)
				&& actual6throwOverallAvgRate.equalsIgnoreCase(exp6throwoverallAvgRate)) {
			
			click(new_CloseBtn);			
			Thread.sleep(1000);
			click(voucherhomeCloseBtn);
			return true;
			
		} else {
			click(new_CloseBtn);			
			Thread.sleep(1000);
			click(voucherhomeCloseBtn);
			return false;

		}
		
		
			
	}
		
		public boolean validteItemQueryStocksInBatchItemAfterSavingSalesReturnsVocher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			Thread.sleep(2000);
			
			//reLogin("inv", "su");
			
			click(inventoryMenu);
			click(InventoryReportsMenu);
			click(itemqueryMenu);
			click(itemTxt);
			Thread.sleep(2000);
			enterText(itemTxt, "BR-BATCH");
			Thread.sleep(3000);
			click(getStockBtn);
			Thread.sleep(3000);

			String actlOpenStockQuantity  = getText(openStockQuantityTxt);
			String expOpenStockQuantity   = "20.000";

			String actlCurrentStock       = getText(currentStockTxt);
			String expCurrentStock        = "3.000";

			String actlavgStockRate       = getText(avgStockRateTxt);
			String expavgStockRate        = "50.1683";

			String actlstockValueTxt      = getText(stockValueTxt);
			String expstockValueTxt       = "150.5050";
			

			System.out.println("=====Stocks In Batch Items======");
			

			System.out.println("actlOpenStockQuantity :" + actlOpenStockQuantity);
			System.out.println("expOpenStockQuantity  :" + expOpenStockQuantity);

			System.out.println("actlCurrentStock      :" + actlCurrentStock);
			System.out.println("expCurrentStock       :" + expCurrentStock);

			System.out.println("actlavgStockRate      :" + actlavgStockRate);
			System.out.println("expavgStockRate       :" + expavgStockRate);

			System.out.println("actlstockValueTxt     :" + actlstockValueTxt);
			System.out.println("expstockValueTxt      :" + expstockValueTxt);
			
			
			ScrollToElement(balanceByWarehouseBtn);
			click(balanceByWarehouseBtn);

			Thread.sleep(3000);

			int count = row1BalanceByWareHouseList.size();

			ArrayList<String> warehouseRow1 = new ArrayList<String>();

			for (int i = 0; i < count; i++) {
				String data = row1BalanceByWareHouseList.get(i).getText();
				warehouseRow1.add(data);
			}

			String actwarehouseRow1List = warehouseRow1.toString();
			String expWarehouseRow1List = "[1, HYD, 2.000, 0.000, 2.000, 60.1010, 120.2020]";

			int count1 = row2BalanceByWareHouseList.size();

			ArrayList<String> warehouseRow2 = new ArrayList<String>();

			for (int i = 0; i < count1; i++) {
				String data = row2BalanceByWareHouseList.get(i).getText();
				warehouseRow2.add(data);
			}

			String actwarehouseRow2List = warehouseRow2.toString();
			String expWarehouseRow2List = "[2, SEC, 1.000, 0.000, 1.000, 30.3030, 30.3030]";
			
			System.out.println(actwarehouseRow1List);
			System.out.println(expWarehouseRow1List);

			System.out.println(actwarehouseRow2List);
			System.out.println(expWarehouseRow2List);
			

			if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
					&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
					&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
					&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)
					&& actwarehouseRow1List.equalsIgnoreCase(expWarehouseRow1List)
					&& actwarehouseRow2List.equalsIgnoreCase(expWarehouseRow2List))
					 {
				return true;
			} else {
				return false;
			}


		}
		
		public boolean validteItemQueryStocksInFifoItemAfterSavingSalesReturnsVocher() throws InterruptedException {

			click(itemTxt);
			Thread.sleep(2000);
			itemTxt.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(2000);
			itemTxt.sendKeys(Keys.BACK_SPACE);
			enterText(itemTxt, "FIFO-BIN");
			Thread.sleep(3000);
			click(getStockBtn);
			Thread.sleep(4000);

			String actlOpenStockQuantity  = getText(openStockQuantityTxt);
			String expOpenStockQuantity   = "20.000";

			String actlCurrentStock       = getText(currentStockTxt);
			String expCurrentStock        = "3.000";

			String actlavgStockRate       = getText(avgStockRateTxt);
			String expavgStockRate        = "166.6667";

			String actlstockValueTxt      = getText(stockValueTxt);
			String expstockValueTxt       = "500.0000";

			System.out.println("=====Stocks In FIFO Items======");

			System.out.println("actlOpenStockQuantity :" + actlOpenStockQuantity);
			System.out.println("expOpenStockQuantity  :" + expOpenStockQuantity);

			System.out.println("actlCurrentStock      :" + actlCurrentStock);
			System.out.println("expCurrentStock       :" + expCurrentStock);

			System.out.println("actlavgStockRate      :" + actlavgStockRate);
			System.out.println("expavgStockRate       :" + expavgStockRate);

			System.out.println("actlstockValueTxt     :" + actlstockValueTxt);
			System.out.println("expstockValueTxt      :" + expstockValueTxt);
			
			ScrollToElement(balanceByWarehouseBtn);
			//click(balanceByWarehouseBtn);

			Thread.sleep(3000);

			int count = row1BalanceByWareHouseList.size();

			ArrayList<String> warehouseRow1 = new ArrayList<String>();

			for (int i = 0; i < count; i++) {
				String data = row1BalanceByWareHouseList.get(i).getText();
				warehouseRow1.add(data);
			}

			String actwarehouseRow1List = warehouseRow1.toString();
			String expWarehouseRow1List = "[1, HYD, 2.000, 0.000, 2.000, 100.0000, 200.0000]";

			int count1 = row2BalanceByWareHouseList.size();

			ArrayList<String> warehouseRow2 = new ArrayList<String>();

			for (int i = 0; i < count1; i++) {
				String data = row2BalanceByWareHouseList.get(i).getText();
				warehouseRow2.add(data);
			}

			String actwarehouseRow2List = warehouseRow2.toString();
			String expWarehouseRow2List = "[2, SEC, 1.000, 0.000, 1.000, 200.0000, 200.0000]";
			
			System.out.println(actwarehouseRow1List);
			System.out.println(expWarehouseRow1List);

			System.out.println(actwarehouseRow2List);
			System.out.println(expWarehouseRow2List);
			

			if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
					&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
					&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
					&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)
					&& actwarehouseRow1List.equalsIgnoreCase(expWarehouseRow1List)
					&& actwarehouseRow2List.equalsIgnoreCase(expWarehouseRow2List))
					 {
				return true;
			} else {
				return false;
			}

		}
		
		
		public boolean validteItemQueryStocksInWAItemAfterSavingSalesReturnsVocher() throws InterruptedException {

			click(itemTxt);
			Thread.sleep(2000);
			itemTxt.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(2000);
			itemTxt.sendKeys(Keys.BACK_SPACE);
			enterText(itemTxt, "WA-RMA");
			Thread.sleep(3000);
			click(getStockBtn);
			Thread.sleep(3000);

			String actlOpenStockQuantity = getText(openStockQuantityTxt);
			String expOpenStockQuantity  = "20.000";

			String actlCurrentStock      = getText(currentStockTxt);
			String expCurrentStock       = "3.000";

			String actlavgStockRate      = getText(avgStockRateTxt);
			String expavgStockRate       = "1,050.3176";

			String actlstockValueTxt     = getText(stockValueTxt);
			String expstockValueTxt      = "3,150.9528";

			ScrollToElement(balanceByWarehouseBtn);
			//click(balanceByWarehouseBtn);

			Thread.sleep(3000);

			int count = row1BalanceByWareHouseList.size();

			ArrayList<String> warehouseRow1 = new ArrayList<String>();

			for (int i = 0; i < count; i++) {
				String data = row1BalanceByWareHouseList.get(i).getText();
				warehouseRow1.add(data);
			}

			String actwarehouseRow1List = warehouseRow1.toString();
			String expWarehouseRow1List = "[1, HYD, 2.000, 0.000, 2.000, 687.9324, 1,375.8648]";

			int count1 = row2BalanceByWareHouseList.size();

			ArrayList<String> warehouseRow2 = new ArrayList<String>();

			for (int i = 0; i < count1; i++) {
				String data = row2BalanceByWareHouseList.get(i).getText();
				warehouseRow2.add(data);
			}

			String actwarehouseRow2List = warehouseRow2.toString();
			String expWarehouseRow2List = "[2, SEC, 1.000, 0.000, 1.000, 1,414.2531, 1,414.2531]";

			
			System.out.println("=====Stocks In WA Items======");

			System.out.println("actlOpenStockQuantity :" + actlOpenStockQuantity);
			System.out.println("expOpenStockQuantity  :" + expOpenStockQuantity);

			System.out.println("actlCurrentStock      :" + actlCurrentStock);
			System.out.println("expCurrentStock       :" + expCurrentStock);

			System.out.println("actlavgStockRate      :" + actlavgStockRate);
			System.out.println("expavgStockRate       :" + expavgStockRate);

			System.out.println("actlstockValueTxt     :" + actlstockValueTxt);
			System.out.println("expstockValueTxt      :" + expstockValueTxt);

			System.out.println("===========Balance By Warehouse=======");

			System.out.println(actwarehouseRow1List);
			System.out.println(expWarehouseRow1List);

			System.out.println(actwarehouseRow2List);
			System.out.println(expWarehouseRow2List);

			

			if (actlOpenStockQuantity.equalsIgnoreCase(expOpenStockQuantity)
					&& actlCurrentStock.equalsIgnoreCase(expCurrentStock)
					&& actlavgStockRate.equalsIgnoreCase(expavgStockRate)
					&& actlstockValueTxt.equalsIgnoreCase(expstockValueTxt)
					&& actwarehouseRow1List.equalsIgnoreCase(expWarehouseRow1List)
					&& actwarehouseRow2List.equalsIgnoreCase(expWarehouseRow2List)) {
				return true;
			} else {
				return false;
			}

		}
		
		@FindBy(xpath = "(//tr[@id='trRender_1'])[1]/td[9]")
		private static WebElement stockLedgerBrBatchAfterSavingSalesReturnsBalanceQty;
		
		@FindBy(xpath = "(//tr[@id='trRender_1'])[1]/td[12]")
		private static WebElement stockLedgerBrBatchAfterSavingSalesReturnsBalanceValue;
		
		@FindBy(xpath = "(//tr[@id='trRender_1'])[1]/td[13]")
		private static WebElement stockLedgerBrBatchAfterSavingSalesReturnsAvgRate;
		
		
		@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[9]")
		private static WebElement stockLedgerBRItemHydBalanceQtyAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[12]")
		private static WebElement stockLedgerBRItemHydBalanceValueAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_3'])[1]/td[13]")
		private static WebElement stockLedgerBRItemHydAvgRateAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_18'])[1]/td[9]")
		private static WebElement stockLedgerBRItemSecBalanceQtyAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_18'])[1]/td[12]")
		private static WebElement stockLedgerBRItemSecBalanceValueAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_18'])[1]/td[13]")
		private static WebElement stockLedgerBRItemSecAvgRateAfterAfterSavingSalesReturns;

	
		
		
		public boolean checkStockLedgerReportForBrBatchItemAfterSavingSalesReturnsVocher()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			click(inventoryMenu);
			click(InventoryReportsMenu);
			click(InventoryStockLedgerMenu);
			Thread.sleep(3000);

			int rowcount = stockLedgerHometableRowCount.size();

			System.out.println(rowcount);

			for (int i = 0; i < rowcount; i++) {
				String actName = ledgerHometableItemNamesList.get(i).getText();

				System.out.println(actName);

				if (actName.equalsIgnoreCase("BR-BATCH")) {
					stockLedgerHometableItemChkboxList.get(i).click();
				}
			}
			
			dropDown(sl_DateOptionDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(2000);

			String actStockLedgerBrBatchBalanceQty   = getText(stockLedgerBrBatchAfterSavingSalesReturnsBalanceQty);
			String expStockLedgerBrBatchBalanceQty   = "3.0000";

			String actStockLedgerBrBatchBalanceValue = getText(stockLedgerBrBatchAfterSavingSalesReturnsBalanceValue);
			String expStockLedgerBrBatchBalanceValue = "150.5050";

			String actStockLedgerBrBatchAvgRate      = getText(stockLedgerBrBatchAfterSavingSalesReturnsAvgRate);
			String expStockLedgerBrBatchAvgRate      = "50.1683";

			System.out.println("actStockLedgerBrBatchBalanceQty        :" + actStockLedgerBrBatchBalanceQty);
			System.out.println("expStockLedgerBrBatchBalanceQty        :" + expStockLedgerBrBatchBalanceQty);

			System.out.println("actStockLedgerBrBatchBalanceValue      :" + actStockLedgerBrBatchBalanceValue);
			System.out.println("expStockLedgerBrBatchBalanceValue      :" + expStockLedgerBrBatchBalanceValue);

			System.out.println("actStockLedgerBrBatchAvgRate      	   :" + actStockLedgerBrBatchAvgRate);
			System.out.println("expStockLedgerBrBatchAvgRate    	   :" + expStockLedgerBrBatchAvgRate);
			
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			click(stockLedgerFilterWarehouseField);
			enterText(stockLedgerFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(stockLedgerLastBtn);
			Thread.sleep(2000);
			
			String actstockLedgerBRItemHydBalanceQtyAfterSavingSalesReturns 		= getText(stockLedgerBRItemHydBalanceQtyAfterSavingSalesReturns);
			String expstockLedgerBRItemHydBalanceQtyAfterSavingSalesReturns 		= "2.0000";

			String actstockLedgerBRItemHydBalanceValueAfterSavingSalesReturns 	    = getText(stockLedgerBRItemHydBalanceValueAfterSavingSalesReturns);
			String expstockLedgerBRItemHydBalanceValueAfterSavingSalesReturns 	    = "120.2020";

			String actstockLedgerBRItemHydAvgRateAfterSavingSalesReturns 		    = getText(stockLedgerBRItemHydAvgRateAfterSavingSalesReturns);
			String expstockLedgerBRItemHydAvgRateAfterSavingSalesReturns 		    = "60.1010";
			
			System.out.println("=========actstockLedger BRItem  HydWareHouse  Values============   : ");
			
			
			System.out.println("actstockLedgerBRItemHydBalanceQtyAfterSavingSalesReturns      :" + actstockLedgerBRItemHydBalanceQtyAfterSavingSalesReturns);
			System.out.println("expstockLedgerBRItemHydBalanceQtyAfterSavingSalesReturns      :" + expstockLedgerBRItemHydBalanceQtyAfterSavingSalesReturns);
			
			System.out.println("actstockLedgerBRItemHydBalanceValueAfterSavingSalesReturns    :" + actstockLedgerBRItemHydBalanceValueAfterSavingSalesReturns);
			System.out.println("expstockLedgerBRItemHydBalanceValueAfterSavingSalesReturns    :" + expstockLedgerBRItemHydBalanceValueAfterSavingSalesReturns);

			System.out.println("actstockLedgerBRItemHydAvgRateAfterSavingSalesReturns         :" + actstockLedgerBRItemHydAvgRateAfterSavingSalesReturns);
			System.out.println("expstockLedgerBRItemHydAvgRateAfterSavingSalesReturns         :" + expstockLedgerBRItemHydAvgRateAfterSavingSalesReturns);		
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockLedgerFilterWarehouseField2);
			Thread.sleep(1000);
			stockLedgerFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockLedgerFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockLedgerFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			Thread.sleep(1000);
			click(stockLedgerLastBtn);
			Thread.sleep(1000);
			click(stockLedgerPreviousBtn);
			
			
			
			String actstockLedgerBRItemSecBalanceQtyAfterSavingSalesReturns 		= getText(stockLedgerBRItemSecBalanceQtyAfterSavingSalesReturns);
			String expstockLedgerBRItemSecBalanceQtyAfterSavingSalesReturns 		= "1.0000";

			String actstockLedgerBRItemSecBalanceValueAfterSavingSalesReturns 	    = getText(stockLedgerBRItemSecBalanceValueAfterSavingSalesReturns);
			String expstockLedgerBRItemSecBalanceValueAfterSavingSalesReturns 	    = "30.3030";

			String actstockLedgerBRItemSecAvgRateAfterAfterSavingSalesReturns 		= getText(stockLedgerBRItemSecAvgRateAfterAfterSavingSalesReturns);
			String expstockLedgerBRItemSecAvgRateAfterAfterSavingSalesReturns 		= "30.3030";
			
			System.out.println("=========actstockLedgerbrItem  SEC WareHouse  Values============   : ");
			
			System.out.println("actstockLedgerBRItemSecBalanceQtyAfterSavingSalesReturns      :" + actstockLedgerBRItemSecBalanceQtyAfterSavingSalesReturns);
			System.out.println("expstockLedgerBRItemSecBalanceQtyAfterSavingSalesReturns      :" + expstockLedgerBRItemSecBalanceQtyAfterSavingSalesReturns);

			System.out.println("actstockLedgerBRItemSecBalanceValueAfterSavingSalesReturns    :" + actstockLedgerBRItemSecBalanceValueAfterSavingSalesReturns);
			System.out.println("expstockLedgerBRItemSecBalanceValueAfterSavingSalesReturns    :" + expstockLedgerBRItemSecBalanceValueAfterSavingSalesReturns);

			System.out.println("actstockLedgerBRItemSecAvgRateAfterAfterSavingSalesReturns    :" + actstockLedgerBRItemSecAvgRateAfterAfterSavingSalesReturns);
			System.out.println("expstockLedgerBRItemSecAvgRateAfterAfterSavingSalesReturns    :" + expstockLedgerBRItemSecAvgRateAfterAfterSavingSalesReturns);
		
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(2000);

			if (actStockLedgerBrBatchBalanceQty.equalsIgnoreCase(expStockLedgerBrBatchBalanceQty)
					&& actStockLedgerBrBatchBalanceValue.equalsIgnoreCase(expStockLedgerBrBatchBalanceValue)
					&& actStockLedgerBrBatchAvgRate.equalsIgnoreCase(expStockLedgerBrBatchAvgRate)
					&& actstockLedgerBRItemHydBalanceQtyAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerBRItemHydBalanceQtyAfterSavingSalesReturns)
					&& actstockLedgerBRItemHydBalanceValueAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerBRItemHydBalanceValueAfterSavingSalesReturns)
					&& actstockLedgerBRItemHydAvgRateAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerBRItemHydAvgRateAfterSavingSalesReturns)
					&& actstockLedgerBRItemSecBalanceQtyAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerBRItemSecBalanceQtyAfterSavingSalesReturns)
					&& actstockLedgerBRItemSecBalanceValueAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerBRItemSecBalanceValueAfterSavingSalesReturns)
					&& actstockLedgerBRItemSecAvgRateAfterAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerBRItemSecAvgRateAfterAfterSavingSalesReturns)) {

				click(report_CloseBtn);
				
				
				return true;
			} else {
				click(report_CloseBtn);
				
				return false;
			}

		}
		
		@FindBy(xpath = "(//tr[@id='trRender_17'])[1]/td[9]")
		private static WebElement stockLedgerFifoItemAfterSavingSalesReturnsBalanceQty;
		
		@FindBy(xpath = "(//tr[@id='trRender_17'])[1]/td[12]")
		private static WebElement stockLedgerFifoItemAfterSavingSalesReturnsBalanceValue;
		
		@FindBy(xpath = "(//tr[@id='trRender_17'])[1]/td[13]")
		private static WebElement stockLedgerFifoItemAfterSavingSalesReturnsAvgRate;
		
		@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[9]")
		private static WebElement stockLedgerFifoItemHydBalanceQtyAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[12]")
		private static WebElement stockLedgerFifoItemHydBalanceValueAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[13]")
		private static WebElement stockLedgerFifoItemHydAvgRateAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[9]")
		private static WebElement stockLedgerFifoItemSecBalanceQtyAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[12]")
		private static WebElement stockLedgerFifoItemSecBalanceValueAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[13]")
		private static WebElement stockLedgerFifoItemSecAvgRateAfterAfterSavingSalesReturns;

		
		
		
		
		public boolean checkStockLedgerReportForFifoItemAfterSavingSalesReturnVocher() // values taken from item query
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

			Thread.sleep(3000);

			if (fifoItemCheckBox.isSelected() == false) {
				click(fifoItemCheckBox);
			}

			if (brBatchItemCheckBox.isSelected() == true) {
				click(brBatchItemCheckBox);
			}

			if (rmaItemCheckBox.isSelected() == true) {
				click(rmaItemCheckBox);

			}

			
			dropDown(sl_DateOptionDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(2000);
 
			String actstockLedgerFifoItemAfterSavingSalesReturnsBalanceQty   = getText(stockLedgerFifoItemAfterSavingSalesReturnsBalanceQty);
			String expstockLedgerFifoItemAfterSavingSalesReturnsBalanceQty   = "3.0000";

			String actstockLedgerFifoItemAfterSavingSalesReturnsBalanceValue = getText(stockLedgerFifoItemAfterSavingSalesReturnsBalanceValue);
			String expstockLedgerFifoItemAfterSavingSalesReturnsBalanceValue = "500.0000";

			String actstockLedgerFifoItemAfterSavingSalesReturnsAvgRate      = getText(stockLedgerFifoItemAfterSavingSalesReturnsAvgRate);
			String expstockLedgerFifoItemAfterSavingSalesReturnsAvgRate      = "166.6667";

			System.out.println("actstockLedgerFifoItemAfterSavingSalesReturnsBalanceQty        :" + actstockLedgerFifoItemAfterSavingSalesReturnsBalanceQty);
			System.out.println("expstockLedgerFifoItemAfterSavingSalesReturnsBalanceQty        :" + expstockLedgerFifoItemAfterSavingSalesReturnsBalanceQty);

			System.out.println("actstockLedgerFifoItemAfterSavingSalesReturnsBalanceValue      :" + actstockLedgerFifoItemAfterSavingSalesReturnsBalanceValue);
			System.out.println("expstockLedgerFifoItemAfterSavingSalesReturnsBalanceValue      :" + expstockLedgerFifoItemAfterSavingSalesReturnsBalanceValue);

			System.out.println("actstockLedgerFifoItemAfterSavingSalesReturnsAvgRate      	   :" + actstockLedgerFifoItemAfterSavingSalesReturnsAvgRate);
			System.out.println("expstockLedgerFifoItemAfterSavingSalesReturnsAvgRate    	   :" + expstockLedgerFifoItemAfterSavingSalesReturnsAvgRate);
			
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			click(stockLedgerFilterWarehouseField);
			enterText(stockLedgerFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(stockLedgerLastBtn);
			Thread.sleep(2000);
			
			String actstockLedgerFifoItemHydBalanceQtyAfterSavingSalesReturns 		= getText(stockLedgerFifoItemHydBalanceQtyAfterSavingSalesReturns);
			String expstockLedgerFifoItemHydBalanceQtyAfterSavingSalesReturns 		= "2.0000";

			String actstockLedgerFifoItemHydBalanceValueAfterSavingSalesReturns 	= getText(stockLedgerFifoItemHydBalanceValueAfterSavingSalesReturns);
			String expstockLedgerFifoItemHydBalanceValueAfterSavingSalesReturns 	= "100.0000";

			String actstockLedgerFifoItemHydAvgRateAfterSavingSalesReturns 		    = getText(stockLedgerFifoItemHydAvgRateAfterSavingSalesReturns);
			String expstockLedgerFifoItemHydAvgRateAfterSavingSalesReturns 		    = "200.0000";
			
			System.out.println("=========actstockLedgerFIFOItem  HydWareHouse  Values============   : ");
			
			System.out.println("actstockLedgerFifoItemHydBalanceQtyAfterSavingSalesReturns      :" + actstockLedgerFifoItemHydBalanceQtyAfterSavingSalesReturns);
			System.out.println("expstockLedgerFifoItemHydBalanceQtyAfterSavingSalesReturns      :" + expstockLedgerFifoItemHydBalanceQtyAfterSavingSalesReturns);
		
			System.out.println("actstockLedgerFifoItemHydBalanceValueAfterSavingSalesReturns    :" + actstockLedgerFifoItemHydBalanceValueAfterSavingSalesReturns);
			System.out.println("expstockLedgerFifoItemHydBalanceValueAfterSavingSalesReturns    :" + expstockLedgerFifoItemHydBalanceValueAfterSavingSalesReturns);

			System.out.println("actstockLedgerFifoItemHydAvgRateAfterSavingSalesReturns         :" + actstockLedgerFifoItemHydAvgRateAfterSavingSalesReturns);
			System.out.println("expstockLedgerFifoItemHydAvgRateAfterSavingSalesReturns         :" + expstockLedgerFifoItemHydAvgRateAfterSavingSalesReturns);		
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockLedgerFilterWarehouseField2);
			Thread.sleep(1000);
			stockLedgerFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockLedgerFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockLedgerFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			Thread.sleep(1000);
			click(stockLedgerLastBtn);
			Thread.sleep(2000);
			
			
			
			String actstockLedgerFifoItemSecBalanceQtyAfterSavingSalesReturns 		= getText(stockLedgerFifoItemSecBalanceQtyAfterSavingSalesReturns);
			String expstockLedgerFifoItemSecBalanceQtyAfterSavingSalesReturns 		= "1.0000";

			String actstockLedgerFifoItemSecBalanceValueAfterSavingSalesReturns 	= getText(stockLedgerFifoItemSecBalanceValueAfterSavingSalesReturns);
			String expstockLedgerFifoItemSecBalanceValueAfterSavingSalesReturns 	= "200.0000";

			String actstockLedgerFifoItemSecAvgRateAfterAfterSavingSalesReturns     = getText(stockLedgerFifoItemSecAvgRateAfterAfterSavingSalesReturns);
			String expstockLedgerFifoItemSecAvgRateAfterAfterSavingSalesReturns 	= "200.0000";
			
			System.out.println("=========actstockLedgerFIFOItem  SEC WareHouse  Values============   : ");
			
			
			
			System.out.println("actstockLedgerFifoItemSecBalanceQtyAfterSavingSalesReturns      :" + actstockLedgerFifoItemSecBalanceQtyAfterSavingSalesReturns);
			System.out.println("expstockLedgerFifoItemSecBalanceQtyAfterSavingSalesReturns      :" + expstockLedgerFifoItemSecBalanceQtyAfterSavingSalesReturns);

			System.out.println("actstockLedgerFifoItemSecBalanceValueAfterSavingSalesReturns    :" + actstockLedgerFifoItemSecBalanceValueAfterSavingSalesReturns);
			System.out.println("expstockLedgerFifoItemSecBalanceValueAfterSavingSalesReturns    :" + expstockLedgerFifoItemSecBalanceValueAfterSavingSalesReturns);

			System.out.println("actstockLedgerFifoItemSecAvgRateAfterAfterSavingSalesReturns    :" + actstockLedgerFifoItemSecAvgRateAfterAfterSavingSalesReturns);
			System.out.println("expstockLedgerFifoItemSecAvgRateAfterAfterSavingSalesReturns    :" + expstockLedgerFifoItemSecAvgRateAfterAfterSavingSalesReturns);
		
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(1000);

			if (actstockLedgerFifoItemAfterSavingSalesReturnsBalanceQty.equalsIgnoreCase(expstockLedgerFifoItemAfterSavingSalesReturnsBalanceQty)
					&& actstockLedgerFifoItemAfterSavingSalesReturnsBalanceValue.equalsIgnoreCase(expstockLedgerFifoItemAfterSavingSalesReturnsBalanceValue)
					&& actstockLedgerFifoItemAfterSavingSalesReturnsAvgRate.equalsIgnoreCase(expstockLedgerFifoItemAfterSavingSalesReturnsAvgRate)
					&& actstockLedgerFifoItemHydBalanceQtyAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerFifoItemHydBalanceQtyAfterSavingSalesReturns)
					&& actstockLedgerFifoItemHydBalanceValueAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerFifoItemHydBalanceValueAfterSavingSalesReturns)
					&& actstockLedgerFifoItemHydAvgRateAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerFifoItemHydAvgRateAfterSavingSalesReturns)
					&& actstockLedgerFifoItemSecBalanceQtyAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerFifoItemSecBalanceQtyAfterSavingSalesReturns)
					&& actstockLedgerFifoItemSecBalanceValueAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerFifoItemSecBalanceValueAfterSavingSalesReturns)
					&& actstockLedgerFifoItemSecAvgRateAfterAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerFifoItemSecAvgRateAfterAfterSavingSalesReturns)) {

				click(report_CloseBtn);
				
				
				return true;
			} else {
				click(report_CloseBtn);
				
				return false;
			}

		}
		
		@FindBy(xpath = "(//tr[@id='trRender_18'])[1]/td[9]")
		private static WebElement stockLedgerRmaItemAfterSavingSalesReturnsBalanceQty;
		
		@FindBy(xpath = "(//tr[@id='trRender_18'])[1]/td[12]")
		private static WebElement stockLedgerRmaItemAfterSavingSalesReturnsBalanceValue;
		
		@FindBy(xpath = "(//tr[@id='trRender_18'])[1]/td[13]")
		private static WebElement stockLedgerRmaItemAfterSavingSalesReturnsAvgRate;
		
		@FindBy(xpath = "(//tr[@id='trRender_10'])[1]/td[9]")
		private static WebElement stockLedgerRmaItemHydBalanceQtyAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_10'])[1]/td[12]")
		private static WebElement stockLedgerRmaItemHydBalanceValueAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_10'])[1]/td[13]")
		private static WebElement stockLedgerRmaItemHydAvgRateAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[9]")
		private static WebElement stockLedgerRmaItemSecBalanceQtyAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[12]")
		private static WebElement stockLedgerRmaItemSecBalanceValueAfterSavingSalesReturns;
		
		@FindBy(xpath = "(//tr[@id='trRender_9'])[1]/td[13]")
		private static WebElement stockLedgerRmaItemSecAvgRateAfterAfterSavingSalesReturns;

		
		
		public boolean checkStockLedgerReportForRmaItemAfterSavingSalesReturnvoucher()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

			Thread.sleep(3000);

			if (fifoItemCheckBox.isSelected() == true) {
				click(fifoItemCheckBox);

			}

			if (brBatchItemCheckBox.isSelected() == true) {
				click(brBatchItemCheckBox);
			}

			if (rmaItemCheckBox.isSelected() == false) {
				click(rmaItemCheckBox);

			}
			
			dropDown(sl_DateOptionDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(2000);
			click(stockLedgerPreviousBtn);

			String actstockLedgerRmaItemAfterSavingSalesReturnsBalanceQty   = getText(stockLedgerRmaItemAfterSavingSalesReturnsBalanceQty);
			String expstockLedgerRmaItemAfterSavingSalesReturnsBalanceQty   = "3.0000";

			String actstockLedgerRmaItemAfterSavingSalesReturnsBalanceValue = getText(stockLedgerRmaItemAfterSavingSalesReturnsBalanceValue);
			String expstockLedgerRmaItemAfterSavingSalesReturnsBalanceValue = "3,150.9528";

			String actstockLedgerRmaItemAfterSavingSalesReturnsAvgRate      = getText(stockLedgerRmaItemAfterSavingSalesReturnsAvgRate);
			String expstockLedgerRmaItemAfterSavingSalesReturnsAvgRate      = "1,050.3176";

			System.out.println("actstockLedgerRmaItemAfterSavingSalesReturnsBalanceQty        :" + actstockLedgerRmaItemAfterSavingSalesReturnsBalanceQty);
			System.out.println("expstockLedgerRmaItemAfterSavingSalesReturnsBalanceQty        :" + expstockLedgerRmaItemAfterSavingSalesReturnsBalanceQty);

			System.out.println("actstockLedgerRmaItemAfterSavingSalesReturnsBalanceValue      :" + actstockLedgerRmaItemAfterSavingSalesReturnsBalanceValue);
			System.out.println("expstockLedgerRmaItemAfterSavingSalesReturnsBalanceValue      :" + expstockLedgerRmaItemAfterSavingSalesReturnsBalanceValue);

			System.out.println("actstockLedgerRmaItemAfterSavingSalesReturnsAvgRate      	  :" + actstockLedgerRmaItemAfterSavingSalesReturnsAvgRate);
			System.out.println("expstockLedgerRmaItemAfterSavingSalesReturnsAvgRate    	      :" + expstockLedgerRmaItemAfterSavingSalesReturnsAvgRate);
			
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			click(stockLedgerFilterWarehouseField);
			enterText(stockLedgerFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(stockLedgerLastBtn);
			Thread.sleep(2000);
			
			String actstockLedgerRmaItemHydBalanceQtyAfterSavingSalesReturns 	= getText(stockLedgerRmaItemHydBalanceQtyAfterSavingSalesReturns);
			String expstockLedgerRmaItemHydBalanceQtyAfterSavingSalesReturns 	= "2.0000";

			String actstockLedgerRmaItemHydBalanceValueAfterSavingSalesReturns 	= getText(stockLedgerRmaItemHydBalanceValueAfterSavingSalesReturns);
			String expstockLedgerRmaItemHydBalanceValueAfterSavingSalesReturns 	= "1,375.8648";

			String actstockLedgerRmaItemHydAvgRateAfterSavingSalesReturns 		= getText(stockLedgerRmaItemHydAvgRateAfterSavingSalesReturns);
			String expstockLedgerRmaItemHydAvgRateAfterSavingSalesReturns 		= "687.9324";
			
			System.out.println("=========actstockLedgerRmaItem  HydWareHouse  Values============   : ");
			
			
			System.out.println("actstockLedgerRmaItemHydBalanceQtyAfterSavingSalesReturns      :" + actstockLedgerRmaItemHydBalanceQtyAfterSavingSalesReturns);
			System.out.println("expstockLedgerRmaItemHydBalanceQtyAfterSavingSalesReturns      :" + expstockLedgerRmaItemHydBalanceQtyAfterSavingSalesReturns);
			
			System.out.println("actstockLedgerRmaItemHydBalanceValueAfterSavingSalesReturns    :" + actstockLedgerRmaItemHydBalanceValueAfterSavingSalesReturns);
			System.out.println("expstockLedgerRmaItemHydBalanceValueAfterSavingSalesReturns    :" + expstockLedgerRmaItemHydBalanceValueAfterSavingSalesReturns);

			System.out.println("actstockLedgerRmaItemHydAvgRateAfterSavingSalesReturns         :" + actstockLedgerRmaItemHydAvgRateAfterSavingSalesReturns);
			System.out.println("expstockLedgerRmaItemHydAvgRateAfterSavingSalesReturns         :" + expstockLedgerRmaItemHydAvgRateAfterSavingSalesReturns);		
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockLedgerFilterWarehouseField2);
			Thread.sleep(1000);
			stockLedgerFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockLedgerFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockLedgerFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			Thread.sleep(1000);
			click(stockLedgerLastBtn);
			Thread.sleep(2000);
			
			
			
			String actstockLedgerRmaItemSecBalanceQtyAfterSavingSalesReturns 		= getText(stockLedgerRmaItemSecBalanceQtyAfterSavingSalesReturns);
			String expstockLedgerRmaItemSecBalanceQtyAfterSavingSalesReturns 		= "1.0000";

			String actstockLedgerRmaItemSecBalanceValueAfterSavingSalesReturns   	= getText(stockLedgerRmaItemSecBalanceValueAfterSavingSalesReturns);
			String expstockLedgerRmaItemSecBalanceValueAfterSavingSalesReturns 	    = "1,414.2531";

			String actstockLedgerRmaItemSecAvgRateAfterAfterSavingSalesReturns 		= getText(stockLedgerRmaItemSecAvgRateAfterAfterSavingSalesReturns);
			String expstockLedgerRmaItemSecAvgRateAfterAfterSavingSalesReturns 		= "1,414.2531";
			
			System.out.println("=========actstockLedgerRmaItem  SEC WareHouse  Values============   : ");
			
			System.out.println("actstockLedgerRmaItemSecBalanceQtyAfterSavingSalesReturns      :" + actstockLedgerRmaItemSecBalanceQtyAfterSavingSalesReturns);
			System.out.println("expstockLedgerRmaItemSecBalanceQtyAfterSavingSalesReturns      :" + expstockLedgerRmaItemSecBalanceQtyAfterSavingSalesReturns);

			System.out.println("actstockLedgerRmaItemSecBalanceValueAfterSavingSalesReturns    :" + actstockLedgerRmaItemSecBalanceValueAfterSavingSalesReturns);
			System.out.println("expstockLedgerRmaItemSecBalanceValueAfterSavingSalesReturns    :" + expstockLedgerRmaItemSecBalanceValueAfterSavingSalesReturns);

			System.out.println("actstockLedgerRmaItemSecAvgRateAfterAfterSavingSalesReturns    :" + actstockLedgerRmaItemSecAvgRateAfterAfterSavingSalesReturns);
			System.out.println("expstockLedgerRmaItemSecAvgRateAfterAfterSavingSalesReturns    :" + expstockLedgerRmaItemSecAvgRateAfterAfterSavingSalesReturns);
		
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(2000);

			if (actstockLedgerRmaItemAfterSavingSalesReturnsBalanceQty.equalsIgnoreCase(expstockLedgerRmaItemAfterSavingSalesReturnsBalanceQty)
					&& actstockLedgerRmaItemAfterSavingSalesReturnsBalanceValue.equalsIgnoreCase(expstockLedgerRmaItemAfterSavingSalesReturnsBalanceValue)
					&& actstockLedgerRmaItemAfterSavingSalesReturnsAvgRate.equalsIgnoreCase(expstockLedgerRmaItemAfterSavingSalesReturnsAvgRate)
					&& actstockLedgerRmaItemHydBalanceQtyAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerRmaItemHydBalanceQtyAfterSavingSalesReturns)
					&& actstockLedgerRmaItemHydBalanceValueAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerRmaItemHydBalanceValueAfterSavingSalesReturns)
					&& actstockLedgerRmaItemHydAvgRateAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerRmaItemHydAvgRateAfterSavingSalesReturns)
					&& actstockLedgerRmaItemSecBalanceQtyAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerRmaItemSecBalanceQtyAfterSavingSalesReturns)
					&& actstockLedgerRmaItemSecBalanceValueAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerRmaItemSecBalanceValueAfterSavingSalesReturns)
					&& actstockLedgerRmaItemSecAvgRateAfterAfterSavingSalesReturns.equalsIgnoreCase(expstockLedgerRmaItemSecAvgRateAfterAfterSavingSalesReturns)) {

				click(report_CloseBtn);
				Thread.sleep(2000);
				click(sl_CloseBtn);
				
				return true;
			} else {
				click(report_CloseBtn);
				Thread.sleep(2000);
				click(sl_CloseBtn);
				return false;
			}

			
			

		}	
		
		@FindBy(xpath = "//a[@id='575']/span")
		private static WebElement InventoryReportsStockMovement;
		
		@FindBy(xpath="//input[@id='liSelectAllMasters']")
		private static WebElement sl_HeaderSelectChkBox;
		
		 // Reports Table Row List  
	  	@FindBy(xpath="(//tr[@id='trRender_0'])[1]/td")
		private static List<WebElement> reportsRow1List;
	  	
	  	@FindBy(xpath="(//tr[@id='trRender_1'])[1]/td")
		private static List<WebElement> reportsRow2List;
		
		@FindBy(xpath="(//tr[@id='trRender_2'])[1]/td")
		private static List<WebElement> reportsRow3List;
		
		@FindBy(xpath="(//tr[@id='trRender_3'])[1]/td")
		private static List<WebElement> reportsRow4List;
		
		@FindBy(xpath="(//tr[@id='trRender_4'])[1]/td")
		private static List<WebElement> reportsRow5List;
		
		@FindBy(xpath="(//tr[@id='trRender_5'])[1]/td")
		private static List<WebElement> reportsRow6List;
		
		@FindBy(xpath="(//tr[@id='trRender_6'])[1]/td")
		private static List<WebElement> reportsRow7List;
		
		@FindBy(xpath="(//tr[@id='trRender_7'])[1]/td")
		private static List<WebElement> reportsRow8List;
		
		@FindBy(xpath="(//tr[@id='trRender_8'])[1]/td")
		private static List<WebElement> reportsRow9List;
		
		@FindBy(xpath="(//tr[@id='trRender_9'])[1]/td")
		private static List<WebElement> reportsRow10List;
		
		@FindBy(xpath="(//tr[@id='trRender_10'])[1]/td")
		private static List<WebElement> reportsRow11List;
		
		@FindBy(xpath="(//tr[@id='trRender_11'])[1]/td")
		private static List<WebElement> reportsRow12List;
		
		@FindBy(xpath="(//tr[@id='trRender_12'])[1]/td")
		private static List<WebElement> reportsRow13List;
		
		@FindBy(xpath="(//tr[@id='trRender_13'])[1]/td")
		private static List<WebElement> reportsRow14List;
		
		@FindBy(xpath="(//tr[@id='trRender_14'])[1]/td")
		private static List<WebElement> reportsRow15List;
		
		@FindBy(xpath="(//tr[@id='trRender_15'])[1]/td")
		private static List<WebElement> reportsRow16List;
		
		@FindBy(xpath="(//tr[@id='trRender_16'])[1]/td")
		private static List<WebElement> reportsRow17List;
		
		@FindBy(xpath="(//tr[@id='trRender_17'])[1]/td")
		private static List<WebElement> reportsRow18List;
		
		

		
		public boolean checkStockMovementReportByDefault()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			
			
			click(inventoryMenu);
			click(InventoryReportsMenu);
			click(InventoryReportsStockMovement);
			click(sl_HeaderSelectChkBox);
			
			dropDown(sl_DateOptionDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[1, WA-RMA, 20.00, 202.1210, 851.00, 41,692.4117, 868.00, -38,743.5798, 3.00, 3,150.9528, 1,050.3176]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[2, FIFO-BIN, 20.00, 202.1210, 850.00, 39,891.7861, 867.00, -39,593.9071, 3.00, 500.0000, 166.6667]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[3, BR-BATCH, 20.00, 202.1210, 850.00, 39,657.3193, 867.00, -39,708.9353, 3.00, 150.5050, 50.1683]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[4, Grand Total, 60.00, 606.3630, 2,551.00, 1,21,241.5171, 2,602.00, -1,18,046.4222, 9.00, 3,801.4578, 1,267.1526]";
			
			
			
		
			
			System.out.println("*********************************checkStockMovementReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			
			
			
			
			
			click(stockMovementFilterWarehouseField);
			enterText(stockMovementFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount1 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount1;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArrayAfterFilter.add(data);
			}
			String actRow1ListAfterFilter = reportsByWarehouseRow1ListArrayAfterFilter.toString();
			String expRow1ListAfterFilter = "[1, WA-RMA, 10.00, 101.0100, 426.00, 15,380.6989, 434.00, -14,105.8441, 2.00, 1,375.8648, 687.9324]";
			
			

			int reportsByWarehouseRow2ListCount1 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount1;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArrayAfterFilter.add(data);
			}
			String actRow2ListAfterFilter = reportsByWarehouseRow2ListArrayAfterFilter.toString();
			String expRow2ListAfterFilter = "[2, FIFO-BIN, 10.00, 101.0100, 425.00, 14,783.0365, 433.00, -14,684.0465, 2.00, 200.0000, 100.0000]";
			
			
			
			int reportsRow3ListCount1 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount1;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArrayAfterFilter.add(data);
			}
			String actRow3ListAfterFilter = reportsRow3ListArrayAfterFilter.toString();
			String expRow3ListAfterFilter = "[3, BR-BATCH, 10.00, 101.0100, 425.00, 14,715.6628, 433.00, -14,696.4708, 2.00, 120.2020, 60.1010]";
			
			
			
			int reportsRow4ListCount1 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount1;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArrayAfterFilter.add(data);
			}
			String actRow4ListAfterFilter = reportsRow4ListArrayAfterFilter.toString();
			String expRow4ListAfterFilter = "[4, Grand Total, 30.00, 303.0300, 1,276.00, 44,879.3982, 1,300.00, -43,486.3614, 6.00, 1,696.0668, 848.0334]";
			
			System.out.println("=========StcokMovement HydWareHouse  Values============   : ");
			
			System.out.println(actRow1ListAfterFilter);
			System.out.println(expRow1ListAfterFilter);
			
			System.out.println(actRow2ListAfterFilter);
			System.out.println(expRow2ListAfterFilter);
			
			System.out.println(actRow3ListAfterFilter);
			System.out.println(expRow3ListAfterFilter);
			
			System.out.println(actRow4ListAfterFilter);
			System.out.println(expRow4ListAfterFilter);
			
			
			
			
			
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockMovementWarehouseField2);
			Thread.sleep(1000);
			stockMovementFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockMovementFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockMovementFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			
			Thread.sleep(2000);
			
			
			
			int reportsByWarehouseRow1ListCount2 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount2;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArraySecLoction.add(data);
			}
			String actRow1ListSecLoction = reportsByWarehouseRow1ListArraySecLoction.toString();
			String expRow1ListSecLoction = "[1, WA-RMA, 10.00, 101.1110, 425.00, 26,311.7128, 434.00, -24,998.5707, 1.00, 1,414.2531, 1,414.2531]";
			
			

			int reportsByWarehouseRow2ListCount2 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount2;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArraySecLoction.add(data);
			}
			String actRow2ListSecLoction = reportsByWarehouseRow2ListArraySecLoction.toString();
			String expRow2ListSecLoction = "[2, FIFO-BIN, 10.00, 101.1110, 425.00, 25,108.7496, 434.00, -25,009.8606, 1.00, 200.0000, 200.0000]";
			
			
			
			int reportsRow3ListCount2 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount2;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArraySecLoction.add(data);
			}
			String actRow3ListSecLoction = reportsRow3ListArraySecLoction.toString();
			String expRow3ListSecLoction = "[3, BR-BATCH, 10.00, 101.1110, 425.00, 24,941.6565, 434.00, -25,012.4645, 1.00, 30.3030, 30.3030]";
			
			
			
			int reportsRow4ListCount2 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount2;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArraySecLoction.add(data);
			}
			String actRow4ListSecLoction = reportsRow4ListArraySecLoction.toString();
			String expRow4ListSecLoction = "[4, Grand Total, 30.00, 303.3330, 1,275.00, 76,362.1189, 1,302.00, -75,020.8958, 3.00, 1,644.5561, 1,644.5561]";
			
			System.out.println("=========StcokMovement  SEC WareHouse  Values============   : ");
			
			
			System.out.println(actRow1ListSecLoction);
			System.out.println(expRow1ListSecLoction);
			
			System.out.println(actRow2ListSecLoction);
			System.out.println(expRow2ListSecLoction);
			
			System.out.println(actRow3ListSecLoction);
			System.out.println(expRow3ListSecLoction);
			
			System.out.println(actRow4ListSecLoction);
			System.out.println(expRow4ListSecLoction);
			
			
		
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(2000);
			
			
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow1ListAfterFilter.equalsIgnoreCase(expRow1ListAfterFilter) && actRow2ListAfterFilter.equalsIgnoreCase(expRow2ListAfterFilter)
				&& actRow3ListAfterFilter.equalsIgnoreCase(expRow3ListAfterFilter) && actRow4ListAfterFilter.equalsIgnoreCase(expRow4ListAfterFilter)
				&& actRow1ListSecLoction.equalsIgnoreCase(expRow1ListSecLoction) && actRow2ListSecLoction.equalsIgnoreCase(expRow2ListSecLoction)
				&& actRow3ListSecLoction.equalsIgnoreCase(expRow3ListSecLoction) && actRow4ListSecLoction.equalsIgnoreCase(expRow4ListSecLoction))
			{
				click(report_CloseBtn);
				return true;
			}
			else
			{
				click(report_CloseBtn);
				return false;
			}
		}
		
		@FindBy(xpath = "//select[@id='RITCombobox__1']")
		private static WebElement stockMovementDropdown;
		
		
		public boolean checkStockMovementReportByWareHouse()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			Thread.sleep(2000);
			
			dropDown(stockMovementDropdown, "1");		
			//Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[1, HYD]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[2, WA-RMA, 10.00, 101.0100, 426.00, 15,380.6989, 434.00, -14,105.8441, 2.00, 1,375.8648, 687.9324]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[3, FIFO-BIN, 10.00, 101.0100, 425.00, 14,783.0365, 433.00, -14,684.0465, 2.00, 200.0000, 100.0000]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[4, BR-BATCH, 10.00, 101.0100, 425.00, 14,715.6628, 433.00, -14,696.4708, 2.00, 120.2020, 60.1010]";
			
			int reportsRow5ListCount = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow5ListCount;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArray.add(data);
			}
			String actRow5List = reportsRow5ListArray.toString();
			String expRow5List = "[5, Sub Total, 30.00, 303.0300, 1,276.00, 44,879.3982, 1,300.00, -43,486.3614, 6.00, 1,696.0668, 848.0334]";
			
			
			int report6thRowListCount = reportsRow6List.size();
			ArrayList<String> report6thRowListArray = new ArrayList<String>();
			for(int i=0;i<report6thRowListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				report6thRowListArray.add(data);
			}
			String actRow6List = report6thRowListArray.toString();
			String expRow6List = "[6, SEC]";
			
			
			int report7thRowListCount = reportsRow7List.size();
			ArrayList<String> report7thRowListArray = new ArrayList<String>();
			for(int i=0;i<report7thRowListCount;i++)
			{
				String data = reportsRow7List.get(i).getText();
				report7thRowListArray.add(data);
			}
			String actRow7List = report7thRowListArray.toString();
			String expRow7List = "[7, WA-RMA, 10.00, 101.1110, 425.00, 26,311.7128, 434.00, -24,998.5707, 1.00, 1,414.2531, 1,414.2531]";
			
			int report8thRowListCount = reportsRow8List.size();
			ArrayList<String> report8thRowListArray = new ArrayList<String>();
			for(int i=0;i<report8thRowListCount;i++)
			{
				String data = reportsRow8List.get(i).getText();
				report8thRowListArray.add(data);
			}
			String actRow8List = report8thRowListArray.toString();
			String expRow8List = "[8, FIFO-BIN, 10.00, 101.1110, 425.00, 25,108.7496, 434.00, -25,009.8606, 1.00, 200.0000, 200.0000]";
			
			
			
			int report9thRowListCount = reportsRow9List.size();
			ArrayList<String> report9thRowListArray = new ArrayList<String>();
			for(int i=0;i<report9thRowListCount;i++)
			{
				String data = reportsRow9List.get(i).getText();
				report9thRowListArray.add(data);
			}
			String actRow9List = report9thRowListArray.toString();
			String expRow9List = "[9, BR-BATCH, 10.00, 101.1110, 425.00, 24,941.6565, 434.00, -25,012.4645, 1.00, 30.3030, 30.3030]";
			
			
			
			int report10thRowListCount = reportsRow10List.size();
			ArrayList<String> report10thRowListArray = new ArrayList<String>();
			for(int i=0;i<report10thRowListCount;i++)
			{
				String data = reportsRow10List.get(i).getText();
				report10thRowListArray.add(data);
			}
			String actRow10List = report10thRowListArray.toString();
			String expRow10List = "[10, Sub Total, 30.00, 303.3330, 1,275.00, 76,362.1189, 1,302.00, -75,020.8958, 3.00, 1,644.5561, 1,644.5561]";
			
			
			
			int reportsRow11ListCount = reportsRow11List.size();
			ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow11ListCount;i++)
			{
				String data = reportsRow11List.get(i).getText();
				reportsRow11ListArray.add(data);
			}
			String actRow11List = reportsRow11ListArray.toString();
			String expRow11List = "[11, Grand Total, 60.00, 606.3630, 2,551.00, 1,21,241.5171, 2,602.00, -1,18,507.2572, 9.00, 3,340.6229, 2,492.5895]";
			
			
			
			
			System.out.println("*********************************checkStockMovementReport BY WareHouse*****************************************");
			
			
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			System.out.println(actRow5List);
			System.out.println(expRow5List);
			
			System.out.println(actRow6List);
			System.out.println(expRow6List);
			
			System.out.println(actRow7List);
			System.out.println(expRow7List);
			
			System.out.println(actRow8List);
			System.out.println(expRow8List);
			
			System.out.println(actRow9List);
			System.out.println(expRow9List);
			
			System.out.println(actRow10List);
			System.out.println(expRow10List);
			
			System.out.println(actRow11List);
			System.out.println(expRow11List);
			
			
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			
			
			
			
			
			click(stockMovementFilterWarehouseField);
			enterText(stockMovementFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount1 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount1;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArrayAfterFilter.add(data);
			}
			String actRow1ListAfterFilter = reportsByWarehouseRow1ListArrayAfterFilter.toString();
			String expRow1ListAfterFilter = "[1, HYD]";
			
			

			int reportsByWarehouseRow2ListCount1 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount1;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArrayAfterFilter.add(data);
			}
			String actRow2ListAfterFilter = reportsByWarehouseRow2ListArrayAfterFilter.toString();
			String expRow2ListAfterFilter = "[2, WA-RMA, 10.00, 101.0100, 426.00, 15,380.6989, 434.00, -14,105.8441, 2.00, 1,375.8648, 687.9324]";
			
			
			
			int reportsRow3ListCount1 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount1;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArrayAfterFilter.add(data);
			}
			String actRow3ListAfterFilter = reportsRow3ListArrayAfterFilter.toString();
			String expRow3ListAfterFilter = "[3, FIFO-BIN, 10.00, 101.0100, 425.00, 14,783.0365, 433.00, -14,684.0465, 2.00, 200.0000, 100.0000]";
			
			
			
			int reportsRow4ListCount1 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount1;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArrayAfterFilter.add(data);
			}
			String actRow4ListAfterFilter = reportsRow4ListArrayAfterFilter.toString();
			String expRow4ListAfterFilter = "[4, BR-BATCH, 10.00, 101.0100, 425.00, 14,715.6628, 433.00, -14,696.4708, 2.00, 120.2020, 60.1010]";
			
			
			int reportsRow5ListCount1 = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow5ListCount1;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArrayAfterFilter.add(data);
			}
			String actRow5ListAfterFilter = reportsRow5ListArrayAfterFilter.toString();
			String expRow5ListAfterFilter = "[5, Grand Total, 30.00, 303.0300, 1,276.00, 44,879.3982, 1,300.00, -43,486.3614, 6.00, 1,696.0668, 848.0334]";
			
			System.out.println("=========StcokMovement HydWareHouse  Values============   : ");
			
			System.out.println(actRow1ListAfterFilter);
			System.out.println(expRow1ListAfterFilter);
			
			System.out.println(actRow2ListAfterFilter);
			System.out.println(expRow2ListAfterFilter);
			
			System.out.println(actRow3ListAfterFilter);
			System.out.println(expRow3ListAfterFilter);
			
			System.out.println(actRow4ListAfterFilter);
			System.out.println(expRow4ListAfterFilter);
			
			System.out.println(actRow5ListAfterFilter);
			System.out.println(expRow5ListAfterFilter);
			
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockMovementWarehouseField2);
			Thread.sleep(1000);
			stockMovementFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockMovementFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockMovementFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			
			Thread.sleep(2000);
			
			
			
			int reportsByWarehouseRow1ListCount2 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount2;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArraySecLoction.add(data);
			}
			String actRow1ListSecLoction = reportsByWarehouseRow1ListArraySecLoction.toString();
			String expRow1ListSecLoction = "[1, SEC]";
			
			

			int reportsByWarehouseRow2ListCount2 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount2;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArraySecLoction.add(data);
			}
			String actRow2ListSecLoction = reportsByWarehouseRow2ListArraySecLoction.toString();
			String expRow2ListSecLoction = "[2, WA-RMA, 10.00, 101.1110, 425.00, 26,311.7128, 434.00, -24,998.5707, 1.00, 1,414.2531, 1,414.2531]";
			
			
			
			int reportsRow3ListCount2 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount2;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArraySecLoction.add(data);
			}
			String actRow3ListSecLoction = reportsRow3ListArraySecLoction.toString();
			String expRow3ListSecLoction = "[3, FIFO-BIN, 10.00, 101.1110, 425.00, 25,108.7496, 434.00, -25,009.8606, 1.00, 200.0000, 200.0000]";
			
			
			
			int reportsRow4ListCount2 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount2;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArraySecLoction.add(data);
			}
			String actRow4ListSecLoction = reportsRow4ListArraySecLoction.toString();
			String expRow4ListSecLoction = "[4, BR-BATCH, 10.00, 101.1110, 425.00, 24,941.6565, 434.00, -25,012.4645, 1.00, 30.3030, 30.3030]";
			
			
			int reportsRow5ListCount2 = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow5ListCount2;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArraySecLoction.add(data);
			}
			String actRow5ListSecLoction = reportsRow5ListArraySecLoction.toString();
			String expRow5ListSecLoction = "[5, Grand Total, 30.00, 303.3330, 1,275.00, 76,362.1189, 1,302.00, -75,020.8958, 3.00, 1,644.5561, 1,644.5561]";
			
			System.out.println("=========StcokMovement  SEC WareHouse  Values============   : ");
			
			
			System.out.println(actRow1ListSecLoction);
			System.out.println(expRow1ListSecLoction);
			
			System.out.println(actRow2ListSecLoction);
			System.out.println(expRow2ListSecLoction);
			
			System.out.println(actRow3ListSecLoction);
			System.out.println(expRow3ListSecLoction);
			
			System.out.println(actRow4ListSecLoction);
			System.out.println(expRow4ListSecLoction);
			
			System.out.println(actRow5ListSecLoction);
			System.out.println(expRow5ListSecLoction);
			
			
		
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(3000);
			
			
			
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List)
					&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
					&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List) && actRow1ListAfterFilter.equalsIgnoreCase(expRow1ListAfterFilter) && actRow2ListAfterFilter.equalsIgnoreCase(expRow2ListAfterFilter)
					&& actRow3ListAfterFilter.equalsIgnoreCase(expRow3ListAfterFilter) && actRow4ListAfterFilter.equalsIgnoreCase(expRow4ListAfterFilter)
					&& actRow5ListAfterFilter.equalsIgnoreCase(expRow5ListAfterFilter) && actRow1ListSecLoction.equalsIgnoreCase(expRow1ListSecLoction) && actRow2ListSecLoction.equalsIgnoreCase(expRow2ListSecLoction)
					&& actRow3ListSecLoction.equalsIgnoreCase(expRow3ListSecLoction) && actRow4ListSecLoction.equalsIgnoreCase(expRow4ListSecLoction)
					&& actRow5ListSecLoction.equalsIgnoreCase(expRow5ListSecLoction))
			{
				click(report_CloseBtn);				
				Thread.sleep(2000);
				click(sl_CloseBtn);				
				return true;
			}
			else
			{
				click(report_CloseBtn);				
				Thread.sleep(2000);
				click(sl_CloseBtn);
				return false;
			}
        }
		
		
		@FindBy(xpath = "//a[@id='578']/span")
		private static WebElement InventoryReportsStockValuation;
		
		public boolean checkStockValuationReportByProduct()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			click(inventoryMenu);
			click(InventoryReportsMenu);
			click(InventoryReportsStockValuation);
			click(sl_HeaderSelectChkBox);
			
			dropDown(sl_DateOptionDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[1, BR-BATCH, BR-BATCH, BR-BATCH, 3.0000, 150.5050, 50.1683, , ]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[2, WA-RMA, WA-RMA, WA-RMA, 3.0000, 3,150.9528, 1,050.3176, , ]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[3, FIFO-BIN, FIFO-BIN, FIFO-BIN, 3.0000, 500.0000, 166.6667, , ]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[4, Grand Total, , , 9.0000, 3,801.4578, 1,267.1526, , ]";
			
			
			
		
			
			System.out.println("*********************************checkStockValuationReport By Product*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			
			
			
			
			
			click(stockValuationFilterWarehouseField);
			enterText(stockValuationFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount1 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount1;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArrayAfterFilter.add(data);
			}
			String actRow1ListAfterFilter = reportsByWarehouseRow1ListArrayAfterFilter.toString();
			String expRow1ListAfterFilter = "[1, WA-RMA, WA-RMA, WA-RMA, 2.0000, 1,375.8648, 687.9324, , ]";
			
			

			int reportsByWarehouseRow2ListCount1 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount1;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArrayAfterFilter.add(data);
			}
			String actRow2ListAfterFilter = reportsByWarehouseRow2ListArrayAfterFilter.toString();
			String expRow2ListAfterFilter = "[2, FIFO-BIN, FIFO-BIN, FIFO-BIN, 2.0000, 200.0000, 100.0000, , ]";
			
			
			
			int reportsRow3ListCount1 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount1;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArrayAfterFilter.add(data);
			}
			String actRow3ListAfterFilter = reportsRow3ListArrayAfterFilter.toString();
			String expRow3ListAfterFilter = "[3, BR-BATCH, BR-BATCH, BR-BATCH, 2.0000, 120.2020, 60.1010, , ]";
			
			
			
			int reportsRow4ListCount1 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount1;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArrayAfterFilter.add(data);
			}
			String actRow4ListAfterFilter = reportsRow4ListArrayAfterFilter.toString();
			String expRow4ListAfterFilter = "[4, Grand Total, , , 6.0000, 1,696.0668, 848.0334, , ]";
			
		
			
			System.out.println("=========StcokMovement HydWareHouse  Values============   : ");
			
			System.out.println(actRow1ListAfterFilter);
			System.out.println(expRow1ListAfterFilter);
			
			System.out.println(actRow2ListAfterFilter);
			System.out.println(expRow2ListAfterFilter);
			
			System.out.println(actRow3ListAfterFilter);
			System.out.println(expRow3ListAfterFilter);
			
			System.out.println(actRow4ListAfterFilter);
			System.out.println(expRow4ListAfterFilter);
			
		
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockValuationWarehouseField2);
			Thread.sleep(1000);
			stockValuationFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockValuationFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockValuationFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			
			Thread.sleep(2000);
			
			
			
			int reportsByWarehouseRow1ListCount2 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount2;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArraySecLoction.add(data);
			}
			String actRow1ListSecLoction = reportsByWarehouseRow1ListArraySecLoction.toString();
			String expRow1ListSecLoction = "[1, WA-RMA, WA-RMA, WA-RMA, 1.0000, 1,414.2531, 1,414.2531, , ]";
			
			

			int reportsByWarehouseRow2ListCount2 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount2;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArraySecLoction.add(data);
			}
			String actRow2ListSecLoction = reportsByWarehouseRow2ListArraySecLoction.toString();
			String expRow2ListSecLoction = "[2, FIFO-BIN, FIFO-BIN, FIFO-BIN, 1.0000, 200.0000, 200.0000, , ]";
			
			
			
			int reportsRow3ListCount2 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount2;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArraySecLoction.add(data);
			}
			String actRow3ListSecLoction = reportsRow3ListArraySecLoction.toString();
			String expRow3ListSecLoction = "[3, BR-BATCH, BR-BATCH, BR-BATCH, 1.0000, 30.3030, 30.3030, , ]";
			
			
			
			int reportsRow4ListCount2 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount2;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArraySecLoction.add(data);
			}
			String actRow4ListSecLoction = reportsRow4ListArraySecLoction.toString();
			String expRow4ListSecLoction = "[4, Grand Total, , , 3.0000, 1,644.5561, 1,644.5561, , ]";
			
			
		
			System.out.println("=========StcokMovement  SEC WareHouse  Values============   : ");
			
			
			System.out.println(actRow1ListSecLoction);
			System.out.println(expRow1ListSecLoction);
			
			System.out.println(actRow2ListSecLoction);
			System.out.println(expRow2ListSecLoction);
			
			System.out.println(actRow3ListSecLoction);
			System.out.println(expRow3ListSecLoction);
			
			System.out.println(actRow4ListSecLoction);
			System.out.println(expRow4ListSecLoction);
			
		
			
		
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(1000);
			
		
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow1ListAfterFilter.equalsIgnoreCase(expRow1ListAfterFilter) && actRow2ListAfterFilter.equalsIgnoreCase(expRow2ListAfterFilter)
				&& actRow3ListAfterFilter.equalsIgnoreCase(expRow3ListAfterFilter) && actRow4ListAfterFilter.equalsIgnoreCase(expRow4ListAfterFilter)
				&& actRow1ListSecLoction.equalsIgnoreCase(expRow1ListSecLoction) && actRow2ListSecLoction.equalsIgnoreCase(expRow2ListSecLoction)
				&& actRow3ListSecLoction.equalsIgnoreCase(expRow3ListSecLoction) && actRow4ListSecLoction.equalsIgnoreCase(expRow4ListSecLoction))
			{
				click(report_CloseBtn);
				return true;
			}
			else
			{
				click(report_CloseBtn);
				return false;
			}
		}
		
		@FindBy(xpath = "//select[@id='RITCombobox__1']")
		private static WebElement stockValuationDropdown;
		
		
		
		public boolean checkStockValuationReportByProductbyInventoryTag()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			Thread.sleep(2000);
			
			dropDown(stockValuationDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(3000);
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[1, FIFO-BIN, FIFO-BIN, FIFO-BIN, 2.0000, 200.0000, 100.0000, HYD, ]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[2, BR-BATCH, BR-BATCH, BR-BATCH, 2.0000, 120.2020, 60.1010, HYD, ]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[3, WA-RMA, WA-RMA, WA-RMA, 2.0000, 1,375.8648, 687.9324, HYD, ]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[4, BR-BATCH, BR-BATCH, BR-BATCH, 1.0000, 30.3030, 30.3030, SEC, ]";
			
			int reportsRow5ListCount = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow5ListCount;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArray.add(data);
			}
			String actRow5List = reportsRow5ListArray.toString();
			String expRow5List = "[5, WA-RMA, WA-RMA, WA-RMA, 1.0000, 1,414.2531, 1,414.2531, SEC, ]";
			
			
			int report6thRowListCount = reportsRow6List.size();
			ArrayList<String> report6thRowListArray = new ArrayList<String>();
			for(int i=0;i<report6thRowListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				report6thRowListArray.add(data);
			}
			String actRow6List = report6thRowListArray.toString();
			String expRow6List = "[6, FIFO-BIN, FIFO-BIN, FIFO-BIN, 1.0000, 200.0000, 200.0000, SEC, ]";
			
			
			int report7thRowListCount = reportsRow7List.size();
			ArrayList<String> report7thRowListArray = new ArrayList<String>();
			for(int i=0;i<report7thRowListCount;i++)
			{
				String data = reportsRow7List.get(i).getText();
				report7thRowListArray.add(data);
			}
			String actRow7List = report7thRowListArray.toString();
			String expRow7List = "[7, Grand Total, , , 9.0000, 3,340.6229, 2,492.5895, , ]";
		
			
			System.out.println("*********************************checkStockValuationReport By ProductbyInventoryTag*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			System.out.println(actRow5List);
			System.out.println(expRow5List);
			
			System.out.println(actRow6List);
			System.out.println(expRow6List);
			
			System.out.println(actRow7List);
			System.out.println(expRow7List);
			
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			
			
			
			
			
			click(stockValuationFilterWarehouseField);
			enterText(stockValuationFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount1 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount1;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArrayAfterFilter.add(data);
			}
			String actRow1ListAfterFilter = reportsByWarehouseRow1ListArrayAfterFilter.toString();
			String expRow1ListAfterFilter = "[1, WA-RMA, WA-RMA, WA-RMA, 2.0000, 1,375.8648, 687.9324, HYD, ]";
			
			

			int reportsByWarehouseRow2ListCount1 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount1;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArrayAfterFilter.add(data);
			}
			String actRow2ListAfterFilter = reportsByWarehouseRow2ListArrayAfterFilter.toString();
			String expRow2ListAfterFilter = "[2, FIFO-BIN, FIFO-BIN, FIFO-BIN, 2.0000, 200.0000, 100.0000, HYD, ]";
			
			
			
			int reportsRow3ListCount1 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount1;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArrayAfterFilter.add(data);
			}
			String actRow3ListAfterFilter = reportsRow3ListArrayAfterFilter.toString();
			String expRow3ListAfterFilter = "[3, BR-BATCH, BR-BATCH, BR-BATCH, 2.0000, 120.2020, 60.1010, HYD, ]";
			
			
			
			int reportsRow4ListCount1 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount1;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArrayAfterFilter.add(data);
			}
			String actRow4ListAfterFilter = reportsRow4ListArrayAfterFilter.toString();
			String expRow4ListAfterFilter = "[4, Grand Total, , , 6.0000, 1,696.0668, 848.0334, , ]";
			

			
			System.out.println("=========StcokMovement HydWareHouse  Values============   : ");
			
			System.out.println(actRow1ListAfterFilter);
			System.out.println(expRow1ListAfterFilter);
			
			System.out.println(actRow2ListAfterFilter);
			System.out.println(expRow2ListAfterFilter);
			
			System.out.println(actRow3ListAfterFilter);
			System.out.println(expRow3ListAfterFilter);
			
			System.out.println(actRow4ListAfterFilter);
			System.out.println(expRow4ListAfterFilter);
			
		
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockValuationWarehouseField2);
			Thread.sleep(1000);
			stockValuationFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockValuationFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockValuationFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			
			Thread.sleep(2000);
			
			
			
			int reportsByWarehouseRow1ListCount2 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount2;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArraySecLoction.add(data);
			}
			String actRow1ListSecLoction = reportsByWarehouseRow1ListArraySecLoction.toString();
			String expRow1ListSecLoction = "[1, WA-RMA, WA-RMA, WA-RMA, 1.0000, 1,414.2531, 1,414.2531, SEC, ]";
			
			

			int reportsByWarehouseRow2ListCount2 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount2;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArraySecLoction.add(data);
			}
			String actRow2ListSecLoction = reportsByWarehouseRow2ListArraySecLoction.toString();
			String expRow2ListSecLoction = "[2, FIFO-BIN, FIFO-BIN, FIFO-BIN, 1.0000, 200.0000, 200.0000, SEC, ]";
			
			
			
			int reportsRow3ListCount2 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount2;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArraySecLoction.add(data);
			}
			String actRow3ListSecLoction = reportsRow3ListArraySecLoction.toString();
			String expRow3ListSecLoction = "[3, BR-BATCH, BR-BATCH, BR-BATCH, 1.0000, 30.3030, 30.3030, SEC, ]";
			
			
			
			int reportsRow4ListCount2 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount2;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArraySecLoction.add(data);
			}
			String actRow4ListSecLoction = reportsRow4ListArraySecLoction.toString();
			String expRow4ListSecLoction = "[4, Grand Total, , , 3.0000, 1,644.5561, 1,644.5561, , ]";
			
			
		
			
			System.out.println("=========StcokMovement  SEC WareHouse  Values============   : ");
			
			
			System.out.println(actRow1ListSecLoction);
			System.out.println(expRow1ListSecLoction);
			
			System.out.println(actRow2ListSecLoction);
			System.out.println(expRow2ListSecLoction);
			
			System.out.println(actRow3ListSecLoction);
			System.out.println(expRow3ListSecLoction);
			
			System.out.println(actRow4ListSecLoction);
			System.out.println(expRow4ListSecLoction);
		
		
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(2000);
			
		
			
			
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List)
					&& actRow7List.equalsIgnoreCase(expRow7List) && actRow1ListAfterFilter.equalsIgnoreCase(expRow1ListAfterFilter) && actRow2ListAfterFilter.equalsIgnoreCase(expRow2ListAfterFilter)
					&& actRow3ListAfterFilter.equalsIgnoreCase(expRow3ListAfterFilter) && actRow4ListAfterFilter.equalsIgnoreCase(expRow4ListAfterFilter)
					&& actRow1ListSecLoction.equalsIgnoreCase(expRow1ListSecLoction) && actRow2ListSecLoction.equalsIgnoreCase(expRow2ListSecLoction)
					&& actRow3ListSecLoction.equalsIgnoreCase(expRow3ListSecLoction) && actRow4ListSecLoction.equalsIgnoreCase(expRow4ListSecLoction))
			{
				click(report_CloseBtn);				
							
				return true;
			}
			else
			{
				click(report_CloseBtn);				
				
				return false;
			}
        }
		
		
		public boolean checkStockValuationReportByInventoryTagByProduct()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			//Thread.sleep(2000);
			
			dropDown(stockValuationDropdown, "2");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(3000);
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[1, HYD]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[2, BR-BATCH, BR-BATCH, BR-BATCH, 2.0000, 120.2020, 60.1010, HYD, ]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[3, FIFO-BIN, FIFO-BIN, FIFO-BIN, 2.0000, 200.0000, 100.0000, HYD, ]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[4, WA-RMA, WA-RMA, WA-RMA, 2.0000, 1,375.8648, 687.9324, HYD, ]";
			
			int reportsRow5ListCount = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow5ListCount;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArray.add(data);
			}
			String actRow5List = reportsRow5ListArray.toString();
			String expRow5List = "[5, Sub Total, , , 6.0000, 1,696.0668, 848.0334, , ]";
			
			
			int report6thRowListCount = reportsRow6List.size();
			ArrayList<String> report6thRowListArray = new ArrayList<String>();
			for(int i=0;i<report6thRowListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				report6thRowListArray.add(data);
			}
			String actRow6List = report6thRowListArray.toString();
			String expRow6List = "[6, SEC]";
			
			
			int report7thRowListCount = reportsRow7List.size();
			ArrayList<String> report7thRowListArray = new ArrayList<String>();
			for(int i=0;i<report7thRowListCount;i++)
			{
				String data = reportsRow7List.get(i).getText();
				report7thRowListArray.add(data);
			}
			String actRow7List = report7thRowListArray.toString();
			String expRow7List = "[7, BR-BATCH, BR-BATCH, BR-BATCH, 1.0000, 30.3030, 30.3030, SEC, ]";
		
			
		
			
			int report8thRowListCount = reportsRow8List.size();
			ArrayList<String> report8thRowListArray = new ArrayList<String>();
			for(int i=0;i<report8thRowListCount;i++)
			{
				String data = reportsRow8List.get(i).getText();
				report8thRowListArray.add(data);
			}
			String actRow8List = report8thRowListArray.toString();
			String expRow8List = "[8, FIFO-BIN, FIFO-BIN, FIFO-BIN, 1.0000, 200.0000, 200.0000, SEC, ]";
			
			
			
			int report9thRowListCount = reportsRow9List.size();
			ArrayList<String> report9thRowListArray = new ArrayList<String>();
			for(int i=0;i<report9thRowListCount;i++)
			{
				String data = reportsRow9List.get(i).getText();
				report9thRowListArray.add(data);
			}
			String actRow9List = report9thRowListArray.toString();
			String expRow9List = "[9, WA-RMA, WA-RMA, WA-RMA, 1.0000, 1,414.2531, 1,414.2531, SEC, ]";
			
			
			
			int report10thRowListCount = reportsRow10List.size();
			ArrayList<String> report10thRowListArray = new ArrayList<String>();
			for(int i=0;i<report10thRowListCount;i++)
			{
				String data = reportsRow10List.get(i).getText();
				report10thRowListArray.add(data);
			}
			String actRow10List = report10thRowListArray.toString();
			String expRow10List = "[10, Sub Total, , , 3.0000, 1,644.5561, 1,644.5561, , ]";
			
			
			
			int reportsRow11ListCount = reportsRow11List.size();
			ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow11ListCount;i++)
			{
				String data = reportsRow11List.get(i).getText();
				reportsRow11ListArray.add(data);
			}
			String actRow11List = reportsRow11ListArray.toString();
			String expRow11List = "[11, Grand Total, , , 9.0000, 3,340.6229, 2,492.5895, , ]";
			
			
			
			System.out.println("*********************************checkStockValuationReport By ProductbyInventoryTag*****************************************");
			
		
			
			
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			System.out.println(actRow5List);
			System.out.println(expRow5List);
			
			System.out.println(actRow6List);
			System.out.println(expRow6List);
			
			System.out.println(actRow7List);
			System.out.println(expRow7List);
			
			System.out.println(actRow8List);
			System.out.println(expRow8List);
			
			System.out.println(actRow9List);
			System.out.println(expRow9List);
			
			System.out.println(actRow10List);
			System.out.println(expRow10List);
			
			System.out.println(actRow11List);
			System.out.println(expRow11List);
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			
			
			
			
			
			click(stockValuationFilterWarehouseField);
			enterText(stockValuationFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount1 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount1;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArrayAfterFilter.add(data);
			}
			String actRow1ListAfterFilter = reportsByWarehouseRow1ListArrayAfterFilter.toString();
			String expRow1ListAfterFilter = "[1, HYD]";
			
			

			int reportsByWarehouseRow2ListCount1 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount1;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArrayAfterFilter.add(data);
			}
			String actRow2ListAfterFilter = reportsByWarehouseRow2ListArrayAfterFilter.toString();
			String expRow2ListAfterFilter = "[2, BR-BATCH, BR-BATCH, BR-BATCH, 2.0000, 120.2020, 60.1010, HYD, ]";
			
			
			
			int reportsRow3ListCount1 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount1;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArrayAfterFilter.add(data);
			}
			String actRow3ListAfterFilter = reportsRow3ListArrayAfterFilter.toString();
			String expRow3ListAfterFilter = "[3, FIFO-BIN, FIFO-BIN, FIFO-BIN, 2.0000, 200.0000, 100.0000, HYD, ]";
			
			
			
			int reportsRow4ListCount1 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount1;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArrayAfterFilter.add(data);
			}
			String actRow4ListAfterFilter = reportsRow4ListArrayAfterFilter.toString();
			String expRow4ListAfterFilter = "[4, WA-RMA, WA-RMA, WA-RMA, 2.0000, 1,375.8648, 687.9324, HYD, ]";
			
			
			int reportsRow5ListCount1 = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow5ListCount1;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArrayAfterFilter.add(data);
			}
			String actRow5ListAfterFilter = reportsRow5ListArrayAfterFilter.toString();
			String expRow5ListAfterFilter = "[5, Grand Total, , , 6.0000, 1,696.0668, 848.0334, , ]";
			
			System.out.println("=========StcokMovement HydWareHouse  Values============   : ");
			
			System.out.println(actRow1ListAfterFilter);
			System.out.println(expRow1ListAfterFilter);
			
			System.out.println(actRow2ListAfterFilter);
			System.out.println(expRow2ListAfterFilter);
			
			System.out.println(actRow3ListAfterFilter);
			System.out.println(expRow3ListAfterFilter);
			
			System.out.println(actRow4ListAfterFilter);
			System.out.println(expRow4ListAfterFilter);
			
			System.out.println(actRow5ListAfterFilter);
			System.out.println(expRow5ListAfterFilter);
			
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockValuationWarehouseField2);
			Thread.sleep(1000);
			stockValuationFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockValuationFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockValuationFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			
			Thread.sleep(2000);
			
			
			
			int reportsByWarehouseRow1ListCount2 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount2;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArraySecLoction.add(data);
			}
			String actRow1ListSecLoction = reportsByWarehouseRow1ListArraySecLoction.toString();
			String expRow1ListSecLoction = "[1, SEC]";
			
			

			int reportsByWarehouseRow2ListCount2 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount2;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArraySecLoction.add(data);
			}
			String actRow2ListSecLoction = reportsByWarehouseRow2ListArraySecLoction.toString();
			String expRow2ListSecLoction = "[2, BR-BATCH, BR-BATCH, BR-BATCH, 1.0000, 30.3030, 30.3030, SEC, ]";
			
			
			
			int reportsRow3ListCount2 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount2;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArraySecLoction.add(data);
			}
			String actRow3ListSecLoction = reportsRow3ListArraySecLoction.toString();
			String expRow3ListSecLoction = "[3, FIFO-BIN, FIFO-BIN, FIFO-BIN, 1.0000, 200.0000, 200.0000, SEC, ]";
			
			
			
			int reportsRow4ListCount2 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount2;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArraySecLoction.add(data);
			}
			String actRow4ListSecLoction = reportsRow4ListArraySecLoction.toString();
			String expRow4ListSecLoction = "[4, WA-RMA, WA-RMA, WA-RMA, 1.0000, 1,414.2531, 1,414.2531, SEC, ]";
			
			
			int reportsRow5ListCount2 = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow5ListCount2;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArraySecLoction.add(data);
			}
			String actRow5ListSecLoction = reportsRow5ListArraySecLoction.toString();
			String expRow5ListSecLoction = "[5, Grand Total, , , 3.0000, 1,644.5561, 1,644.5561, , ]";
			
			System.out.println("=========StcokMovement  SEC WareHouse  Values============   : ");
			
			
			System.out.println(actRow1ListSecLoction);
			System.out.println(expRow1ListSecLoction);
			
			System.out.println(actRow2ListSecLoction);
			System.out.println(expRow2ListSecLoction);
			
			System.out.println(actRow3ListSecLoction);
			System.out.println(expRow3ListSecLoction);
			
			System.out.println(actRow4ListSecLoction);
			System.out.println(expRow4ListSecLoction);
			
			System.out.println(actRow5ListSecLoction);
			System.out.println(expRow5ListSecLoction);
			
			
		
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(3000);
			
			
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List)
					&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
					&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List) && actRow1ListAfterFilter.equalsIgnoreCase(expRow1ListAfterFilter) && actRow2ListAfterFilter.equalsIgnoreCase(expRow2ListAfterFilter)
					&& actRow3ListAfterFilter.equalsIgnoreCase(expRow3ListAfterFilter) && actRow4ListAfterFilter.equalsIgnoreCase(expRow4ListAfterFilter)
					&& actRow5ListAfterFilter.equalsIgnoreCase(expRow5ListAfterFilter) && actRow1ListSecLoction.equalsIgnoreCase(expRow1ListSecLoction) && actRow2ListSecLoction.equalsIgnoreCase(expRow2ListSecLoction)
					&& actRow3ListSecLoction.equalsIgnoreCase(expRow3ListSecLoction) && actRow4ListSecLoction.equalsIgnoreCase(expRow4ListSecLoction)
					&& actRow5ListSecLoction.equalsIgnoreCase(expRow5ListSecLoction))
			{
				click(report_CloseBtn);				
				Thread.sleep(2000);
				click(sl_CloseBtn);				
				return true;
			}
			else
			{
				click(report_CloseBtn);				
				Thread.sleep(2000);
				click(sl_CloseBtn);
				return false;
			}
        
		
		
		}	
		
		@FindBy(xpath = "//a[@id='621']/span")
		private static WebElement InventoryReportsStockAgeingAnalysisMenu;
		
		@FindBy(xpath = "//a[@id='582']/span")
		private static WebElement InventoryReportsAgeingAnalysis;
		
		public boolean checkStockAgeingAnalysisReportByDefault()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			click(inventoryMenu);
			click(InventoryReportsMenu);
			click(InventoryReportsStockAgeingAnalysisMenu);
			click(InventoryReportsAgeingAnalysis);
			click(sl_HeaderSelectChkBox);
			
			dropDown(sl_DateOptionDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[1, WA-RMA, 209.00, 3,150.9528]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[2, FIFO-BIN, 209.00, 500.0000]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[3, BR-BATCH, 209.00, 150.5050]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[4, Grand Total, 627.00, 3,801.4578]";
			
			
			
		
			
			System.out.println("*********************************checkStockAgeing AnalysisReport By Default*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			
			
			
			click(stockAgeingFilterWarehouseField);
			enterText(stockAgeingFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount1 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount1;i++)
			{
				String data = reportsRow1List.get(i).getText();
				if(data.isEmpty()==false)
				{
					reportsByWarehouseRow1ListArrayAfterFilter.add(data);
				}
				
			}
			String actRow1ListAfterFilter = reportsByWarehouseRow1ListArrayAfterFilter.toString();
			String expRow1ListAfterFilter = "[1, WA-RMA, 106.00, 1,375.8648, 12.00, 8,255.19, 94.00, 64,665.64]";
			
			

			int reportsByWarehouseRow2ListCount1 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount1;i++)
			{
				String data = reportsRow2List.get(i).getText();
				if(data.isEmpty()==false)
				{
					reportsByWarehouseRow2ListArrayAfterFilter.add(data);
				}
				
			}
			String actRow2ListAfterFilter = reportsByWarehouseRow2ListArrayAfterFilter.toString();
			String expRow2ListAfterFilter = "[2, FIFO-BIN, 106.00, 200.0000, 12.00, 1,200.00, 94.00, 9,400.00]";
			
			
			
			int reportsRow3ListCount1 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount1;i++)
			{
				String data = reportsRow3List.get(i).getText();
				if(data.isEmpty()==false)
				{
					reportsRow3ListArrayAfterFilter.add(data);
				}
				
			}
			String actRow3ListAfterFilter = reportsRow3ListArrayAfterFilter.toString();
			String expRow3ListAfterFilter = "[3, BR-BATCH, 106.00, 120.2020, 12.00, 721.21, 94.00, 5,649.49]";
			
			
			
			int reportsRow4ListCount1 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArrayAfterFilter = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount1;i++)
			{
				String data = reportsRow4List.get(i).getText();
				if(data.isEmpty()==false)
				{
					reportsRow4ListArrayAfterFilter.add(data);
				}
				
			}
			String actRow4ListAfterFilter = reportsRow4ListArrayAfterFilter.toString();
			String expRow4ListAfterFilter = "[4, Grand Total, 318.00, 1,696.0668, 36.00, 10,176.40, 282.00, 79,715.14]";
			

			
			System.out.println("=========StcokMovement HydWareHouse  Values============   : ");
			
			System.out.println(actRow1ListAfterFilter);
			System.out.println(expRow1ListAfterFilter);
			
			System.out.println(actRow2ListAfterFilter);
			System.out.println(expRow2ListAfterFilter);
			
			System.out.println(actRow3ListAfterFilter);
			System.out.println(expRow3ListAfterFilter);
			
			System.out.println(actRow4ListAfterFilter);
			System.out.println(expRow4ListAfterFilter);
			
		
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockAgeingWarehouseField2);
			Thread.sleep(1000);
			stockAgeingFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockAgeingFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockAgeingFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			
			Thread.sleep(2000);
			
			
			
			int reportsByWarehouseRow1ListCount2 = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount2;i++)
			{
				String data = reportsRow1List.get(i).getText();
				if(data.isEmpty()==false)
				{
					reportsByWarehouseRow1ListArraySecLoction.add(data);
				}
				
			}
			String actRow1ListSecLoction = reportsByWarehouseRow1ListArraySecLoction.toString();
			String expRow1ListSecLoction = "[1, WA-RMA, 103.00, 1,414.2531, 12.00, 16,971.04, 91.00, 1,28,697.03]";
			
			

			int reportsByWarehouseRow2ListCount2 = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount2;i++)
			{
				String data = reportsRow2List.get(i).getText();
				if(data.isEmpty()==false)
				{
					reportsByWarehouseRow2ListArraySecLoction.add(data);
				}
				
			}
			String actRow2ListSecLoction = reportsByWarehouseRow2ListArraySecLoction.toString();
			String expRow2ListSecLoction = "[2, FIFO-BIN, 103.00, 200.0000, 12.00, 2,400.00, 91.00, 18,200.00]";
			
			
			
			int reportsRow3ListCount2 = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount2;i++)
			{
				String data = reportsRow3List.get(i).getText();
				if(data.isEmpty()==false)
				{
					reportsRow3ListArraySecLoction.add(data);
				}
				
			}
			String actRow3ListSecLoction = reportsRow3ListArraySecLoction.toString();
			String expRow3ListSecLoction = "[3, BR-BATCH, 103.00, 30.3030, 12.00, 363.64, 91.00, 2,757.57]";
			
			
			
			int reportsRow4ListCount2 = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArraySecLoction = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount2;i++)
			{
				String data = reportsRow4List.get(i).getText();
				if(data.isEmpty()==false)
				{
					reportsRow4ListArraySecLoction.add(data);
				}
				
			}
			String actRow4ListSecLoction = reportsRow4ListArraySecLoction.toString();
			String expRow4ListSecLoction = "[4, Grand Total, 309.00, 1,644.5561, 36.00, 19,734.67, 273.00, 1,49,654.61]";
			
			
		
			
			System.out.println("=========StcokMovement  SEC WareHouse  Values============   : ");
			
			
			System.out.println(actRow1ListSecLoction);
			System.out.println(expRow1ListSecLoction);
			
			System.out.println(actRow2ListSecLoction);
			System.out.println(expRow2ListSecLoction);
			
			System.out.println(actRow3ListSecLoction);
			System.out.println(expRow3ListSecLoction);
			
			System.out.println(actRow4ListSecLoction);
			System.out.println(expRow4ListSecLoction);
		
		
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(3000);
			
			
		
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow1ListAfterFilter.equalsIgnoreCase(expRow1ListAfterFilter) && actRow2ListAfterFilter.equalsIgnoreCase(expRow2ListAfterFilter)
				&& actRow3ListAfterFilter.equalsIgnoreCase(expRow3ListAfterFilter) && actRow4ListAfterFilter.equalsIgnoreCase(expRow4ListAfterFilter)
				&& actRow1ListSecLoction.equalsIgnoreCase(expRow1ListSecLoction) && actRow2ListSecLoction.equalsIgnoreCase(expRow2ListSecLoction)
				&& actRow3ListSecLoction.equalsIgnoreCase(expRow3ListSecLoction) && actRow4ListSecLoction.equalsIgnoreCase(expRow4ListSecLoction))
			{
				click(report_CloseBtn);
				return true;
			}
			else
			{
				click(report_CloseBtn);
				return false;
			}
		}

		@FindBy(xpath = "//select[@id='RITCombobox__1']")
		private static WebElement stockAgeingDropdown;
		
		
		public boolean checkStockAgeingAnalysisReportByWarehouse()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			Thread.sleep(2000);
			
			dropDown(stockAgeingDropdown, "1");		
			Thread.sleep(3000);
			click(sl_OkBtn);
			Thread.sleep(3000);
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[1, HYD]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[2, WA-RMA, 106.00, 1,375.8648]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[3, FIFO-BIN, 106.00, 200.0000]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[4, BR-BATCH, 106.00, 120.2020]";
			
			int reportsRow5ListCount = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArray.add(data);
			}
			String actRow5List = reportsRow5ListArray.toString();
			String expRow5List = "[5, Sub Total, 318.00, 1,696.0668]";
			
			
			int report6thRowListCount = reportsRow6List.size();
			ArrayList<String> report6thRowListArray = new ArrayList<String>();
			for(int i=0;i<report6thRowListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				report6thRowListArray.add(data);
			}
			String actRow6List = report6thRowListArray.toString();
			String expRow6List = "[6, SEC]";
			
			
			int report7thRowListCount = reportsRow7List.size();
			ArrayList<String> report7thRowListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow7List.get(i).getText();
				report7thRowListArray.add(data);
			}
			String actRow7List = report7thRowListArray.toString();
			String expRow7List = "[7, WA-RMA, 103.00, 1,414.2531]";
		
			
		
			
			int report8thRowListCount = reportsRow8List.size();
			ArrayList<String> report8thRowListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow8List.get(i).getText();
				report8thRowListArray.add(data);
			}
			String actRow8List = report8thRowListArray.toString();
			String expRow8List = "[8, FIFO-BIN, 103.00, 200.0000]";
			
			
			
			int report9thRowListCount = reportsRow9List.size();
			ArrayList<String> report9thRowListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow9List.get(i).getText();
				report9thRowListArray.add(data);
			}
			String actRow9List = report9thRowListArray.toString();
			String expRow9List = "[9, BR-BATCH, 103.00, 30.3030]";
			
			
			
			int report10thRowListCount = reportsRow10List.size();
			ArrayList<String> report10thRowListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow10List.get(i).getText();
				report10thRowListArray.add(data);
			}
			String actRow10List = report10thRowListArray.toString();
			String expRow10List = "[10, Sub Total, 309.00, 1,644.5561]";
			
			
			
			int reportsRow11ListCount = reportsRow11List.size();
			ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow11List.get(i).getText();
				reportsRow11ListArray.add(data);
			}
			String actRow11List = reportsRow11ListArray.toString();
			String expRow11List = "[11, Grand Total, 627.00, 3,340.6229]";
			
			
			
			System.out.println("*********************************checkStock Ageing Analysis Report By WareHouse*****************************************");
			
		
			
			
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			System.out.println(actRow5List);
			System.out.println(expRow5List);
			
			System.out.println(actRow6List);
			System.out.println(expRow6List);
			
			System.out.println(actRow7List);
			System.out.println(expRow7List);
			
			System.out.println(actRow8List);
			System.out.println(expRow8List);
			
			System.out.println(actRow9List);
			System.out.println(expRow9List);
			
			System.out.println(actRow10List);
			System.out.println(expRow10List);
			
			System.out.println(actRow11List);
			System.out.println(expRow11List);
			
			
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List)
					&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
					&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List))
			{
				click(report_CloseBtn);				
							
				return true;
			}
			else
			{
				click(report_CloseBtn);				
				
				return false;
			}
        
		
		
		}	
		
		@FindBy(xpath = "//input[@id='MasterGroup__101']")
		private static WebElement stockAgeingAnalysisWarehouseField;
		
		public boolean checkStockAgeingAnalysisReportByWarehouseHyd()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			Thread.sleep(2000);
			click(stockAgeingAnalysisWarehouseField);
			enterText(stockAgeingAnalysisWarehouseField, "HYD");
			Thread.sleep(2000);
			dropDown(stockAgeingDropdown, 0);
			
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(3000);
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[1, WA-RMA, 106.00, 1,375.8648]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[2, FIFO-BIN, 106.00, 200.0000]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[3, BR-BATCH, 106.00, 120.2020]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[4, Grand Total, 318.00, 1,696.0668]";
			
		

			System.out.println("*********************************checkStock Ageing Analysis Report By WareHouse HYD*****************************************");
			
		
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
		
			
		
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
					&& actRow4List.equalsIgnoreCase(expRow4List))
			{
				click(report_CloseBtn);				
							
				return true;
			}
			else
			{
				click(report_CloseBtn);				
				
				return false;
			}
       
		}	
		
		
		public boolean checkStockAgeingAnalysisReportByWarehouseSEC()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			Thread.sleep(2000);
			click(stockAgeingAnalysisWarehouseField);			
			stockAgeingAnalysisWarehouseField.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(2000);
			stockAgeingAnalysisWarehouseField.sendKeys(Keys.BACK_SPACE);			
			enterText(stockAgeingAnalysisWarehouseField, "SEC");			
			Thread.sleep(4000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[1, WA-RMA, 103.00, 1,414.2531]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[2, FIFO-BIN, 103.00, 200.0000]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[3, BR-BATCH, 103.00, 30.3030]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=0;i<4;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[4, Grand Total, 309.00, 1,644.5561]";
			
			
			

			System.out.println("*********************************checkStock Ageing Analysis Report By WareHouse SEC*****************************************");
			
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
		
		
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
					&& actRow4List.equalsIgnoreCase(expRow4List))
			{
				click(report_CloseBtn);				
				Thread.sleep(2000);
				click(sl_CloseBtn);		
				return true;
			}
			else
			{
			   click(report_CloseBtn);
			   Thread.sleep(2000);
			   click(sl_CloseBtn);
			   return false;
			}
        
		
		
		}	
		
		@FindBy(xpath = "//a[@id='645']/span")
		private static WebElement InventoryReportsStockBalanceByWarehouseMenu;
		
		
		public boolean checkStockBalanceByWarehouseReport()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			
			click(inventoryMenu);
			click(InventoryReportsMenu);
			Thread.sleep(2000);
			moveToElement(InventoryReportsStockBalanceByWarehouseMenu);
			click(sl_HeaderSelectChkBox);
			
			dropDown(sl_DateOptionDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[1, HYD HYD]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[2, FIFO-BIN, 0, 0, 2.0000000, , 2.0000000, 100.00, 166.6667]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[3, BR-BATCH, 0, 0, 2.0000000, , 2.0000000, 60.10, 50.1683]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[4, WA-RMA, 0, 0, 2.0000000, , 2.0000000, 687.93, 1,050.3176]";
			
		
			int reportsRow5ListCount = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow5ListCount;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArray.add(data);
			}
			String actRow5List = reportsRow5ListArray.toString();
			String expRow5List = "[5, Sub Total, , , 6.0000000, , 6.0000000, 848.03, 1,267.1526]";
			
			
			int report6thRowListCount = reportsRow6List.size();
			ArrayList<String> report6thRowListArray = new ArrayList<String>();
			for(int i=0;i<report6thRowListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				report6thRowListArray.add(data);
			}
			String actRow6List = report6thRowListArray.toString();
			String expRow6List = "[6, SEC SEC]";
			
			
			int report7thRowListCount = reportsRow7List.size();
			ArrayList<String> report7thRowListArray = new ArrayList<String>();
			for(int i=0;i<report7thRowListCount;i++)
			{
				String data = reportsRow7List.get(i).getText();
				report7thRowListArray.add(data);
			}
			String actRow7List = report7thRowListArray.toString();
			String expRow7List = "[7, BR-BATCH, 0, 0, 1.0000000, , 1.0000000, 60.10, 50.1683]";
		
			
		
			
			int report8thRowListCount = reportsRow8List.size();
			ArrayList<String> report8thRowListArray = new ArrayList<String>();
			for(int i=0;i<report8thRowListCount;i++)
			{
				String data = reportsRow8List.get(i).getText();
				report8thRowListArray.add(data);
			}
			String actRow8List = report8thRowListArray.toString();
			String expRow8List = "[8, FIFO-BIN, 0, 0, 1.0000000, , 1.0000000, 100.00, 166.6667]";
			
			
			
			int report9thRowListCount = reportsRow9List.size();
			ArrayList<String> report9thRowListArray = new ArrayList<String>();
			for(int i=0;i<report9thRowListCount;i++)
			{
				String data = reportsRow9List.get(i).getText();
				report9thRowListArray.add(data);
			}
			String actRow9List = report9thRowListArray.toString();
			String expRow9List = "[9, WA-RMA, 0, 0, 1.0000000, , 1.0000000, 687.93, 1,050.3176]";
			
			
			
			int report10thRowListCount = reportsRow10List.size();
			ArrayList<String> report10thRowListArray = new ArrayList<String>();
			for(int i=0;i<report10thRowListCount;i++)
			{
				String data = reportsRow10List.get(i).getText();
				report10thRowListArray.add(data);
			}
			String actRow10List = report10thRowListArray.toString();
			String expRow10List = "[10, Sub Total, , , 3.0000000, , 3.0000000, 848.03, 1,267.1526]";
			
			int reportsRow11ListCount = reportsRow11List.size();
			ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
			for(int i=0;i<reportsRow11ListCount;i++)
			{
				String data = reportsRow11List.get(i).getText();
				reportsRow11ListArray.add(data);
			}
			String actRow11List = reportsRow11ListArray.toString();
			String expRow11List = "[11, Grand Total, , , 9.0000000, , 9.0000000, 1,696.07, 2,534.3052]";
			
			
			
			System.out.println("*********************************checkStockBalanceByWarehouseReport*****************************************");
			
		
			
			
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			System.out.println(actRow5List);
			System.out.println(expRow5List);
			
			System.out.println(actRow6List);
			System.out.println(expRow6List);
			
			System.out.println(actRow7List);
			System.out.println(expRow7List);
			
			System.out.println(actRow8List);
			System.out.println(expRow8List);
			
			System.out.println(actRow9List);
			System.out.println(expRow9List);
			
			System.out.println(actRow10List);
			System.out.println(expRow10List);
			
			System.out.println(actRow11List);
			System.out.println(expRow11List);
			
			
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List)
					&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
					&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List))
			{
				click(report_CloseBtn);				
				Thread.sleep(2000);
				click(sl_CloseBtn);		
				return true;
			}
			else
			{
				click(report_CloseBtn);				
				Thread.sleep(2000);
				click(sl_CloseBtn);		
				return false;
			}
        
		
		
		}	
		
		@FindBy(xpath = "//a[@id='557']/span")
		private static WebElement InventoryStockStatementMenu;
		
		public boolean checkStockStatementReporFortBrBatchItem()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			
			
			click(inventoryMenu);
			click(InventoryReportsMenu);
			click(InventoryStockStatementMenu);
			Thread.sleep(3000);

			int rowcount = stockLedgerHometableRowCount.size();

			System.out.println(rowcount);

			for (int i = 0; i < rowcount; i++) {
				String actName = ledgerHometableItemNamesList.get(i).getText();

				System.out.println(actName);

				if (actName.equalsIgnoreCase("BR-BATCH")) {
					stockLedgerHometableItemChkboxList.get(i).click();
				}
			}
			
			dropDown(sl_DateOptionDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(23000);
			
			int reportsByWarehouseRow13ListCount = reportsRow13List.size();
			ArrayList<String> reportsByWarehouseRow13ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow13ListCount;i++)
			{
				String data = reportsRow13List.get(i).getText();
				reportsByWarehouseRow13ListArray.add(data);
			}
			String actRow13List = reportsByWarehouseRow13ListArray.toString();
			String expRow13List = "[184, 07/09/2021, 1.0000, 2.0000, , 3.0000, 150.5050, BR-BATCH]";
			
			int reportsByWarehouseRow14ListCount = reportsRow14List.size();
			ArrayList<String> reportsByWarehouseRow14ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow14ListCount;i++)
			{
				String data = reportsRow14List.get(i).getText();
				reportsByWarehouseRow14ListArray.add(data);
			}
			String actRow14List = reportsByWarehouseRow14ListArray.toString();
			String expRow14List = "[185, Total, 2,414.0000, 850.0000, 867.0000, 3.0000, 150.5050, ]";
			
			System.out.println("===========Stock Statement Report for Bt Item====================");
			
			System.out.println(actRow13List);
			System.out.println(expRow13List);
			
			System.out.println(actRow14List);
			System.out.println(expRow14List);
			
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			
			click(stockStatementFilterWarehouseField);
			enterText(stockStatementFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(23000);
			
			
			int reportsByWarehouseRow14ListCount1 = reportsRow14List.size();
			ArrayList<String> reportsByWarehouseRow14ListArray1 = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow14ListCount1;i++)
			{
				String data = reportsRow14List.get(i).getText();
				reportsByWarehouseRow14ListArray1.add(data);
			}
			String actRow14List1 = reportsByWarehouseRow14ListArray1.toString();
			String expRow14List1 = "[166, 07/09/2021, 1.0000, 1.0000, , 2.0000, 120.2020, BR-BATCH]";
			
			int reportsByWarehouseRow15ListCount = reportsRow15List.size();
			ArrayList<String> reportsByWarehouseRow15ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow15ListCount;i++)
			{
				String data = reportsRow15List.get(i).getText();
				reportsByWarehouseRow15ListArray.add(data);
			}
			String actRow15List = reportsByWarehouseRow15ListArray.toString();
			String expRow15List = "[167, Total, 870.0000, 425.0000, 433.0000, 2.0000, 120.2020, ]";
			
			System.out.println("===========Stock Statement Report for Bt Item HYD Location====================");
			
			System.out.println(actRow14List1);
			System.out.println(expRow14List1);
			
			System.out.println(actRow15List);
			System.out.println(expRow15List);
			
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockStatementWarehouseField2);
			Thread.sleep(1000);
			stockStatementFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockStatementFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockStatementFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(23000);
			
			int reportsByWarehouseRow17ListCount = reportsRow17List.size();
			ArrayList<String> reportsByWarehouseRow17ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow17ListCount;i++)
			{
				String data = reportsRow17List.get(i).getText();
				reportsByWarehouseRow17ListArray.add(data);
			}
			String actRow17List = reportsByWarehouseRow17ListArray.toString();
			String expRow17List = "[169, 07/09/2021, , 1.0000, , 1.0000, 30.3030, BR-BATCH]";
			
			int reportsByWarehouseRow18ListCount = reportsRow18List.size();
			ArrayList<String> reportsByWarehouseRow18ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow18ListCount;i++)
			{
				String data = reportsRow18List.get(i).getText();
				reportsByWarehouseRow18ListArray.add(data);
			}
			String actRow18List = reportsByWarehouseRow18ListArray.toString();
			String expRow18List = "[170, Total, 1,273.0000, 425.0000, 434.0000, 1.0000, 30.3030, ]";
			
			System.out.println("===========Stock Statement Report for Bt Item SCE Location====================");
			
			System.out.println(actRow17List);
			System.out.println(expRow17List);
			
			System.out.println(actRow18List);
			System.out.println(expRow18List);
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(3000);
			
			
			if(actRow13List.equalsIgnoreCase(expRow13List) && actRow14List.equalsIgnoreCase(expRow14List)
					&& actRow14List1.equalsIgnoreCase(expRow14List1) && actRow15List.equalsIgnoreCase(expRow15List)
					&& actRow17List.equalsIgnoreCase(expRow17List) && actRow18List.equalsIgnoreCase(expRow18List))
			{
				click(report_CloseBtn);				
				return true;
			}
			else
			{
				click(report_CloseBtn);							
				return false;
			}
			
			
			
			
			
		}
		
		public boolean checkStockStatementReportForFifoItem() 
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

			Thread.sleep(3000);

			if (fifoItemCheckBox.isSelected() == false) {
				click(fifoItemCheckBox);
			}

			if (brBatchItemCheckBox.isSelected() == true) {
				click(brBatchItemCheckBox);
			}

			if (rmaItemCheckBox.isSelected() == true) {
				click(rmaItemCheckBox);

			}

			
			dropDown(sl_DateOptionDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(23000);
			
			int reportsByWarehouseRow13ListCount = reportsRow13List.size();
			ArrayList<String> reportsByWarehouseRow13ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow13ListCount;i++)
			{
				String data = reportsRow13List.get(i).getText();
				reportsByWarehouseRow13ListArray.add(data);
			}
			String actRow13List = reportsByWarehouseRow13ListArray.toString();
			String expRow13List = "[184, 07/09/2021, 1.0000, 2.0000, , 3.0000, 500.0000, FIFO-BIN]";
			
			
			
			int reportsByWarehouseRow14ListCount = reportsRow14List.size();
			ArrayList<String> reportsByWarehouseRow14ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow14ListCount;i++)
			{
				String data = reportsRow14List.get(i).getText();
				reportsByWarehouseRow14ListArray.add(data);
			}
			String actRow14List = reportsByWarehouseRow14ListArray.toString();
			String expRow14List = "[185, Total, 2,416.0000, 850.0000, 867.0000, 3.0000, 500.0000, ]";
			
			System.out.println("===========Stock Statement Report for FIFO Item ====================");
			
			System.out.println(actRow13List);
			System.out.println(expRow13List);
			
			System.out.println(actRow14List);
			System.out.println(expRow14List);
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			
			click(stockStatementFilterWarehouseField);
			enterText(stockStatementFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(23000);
			
			
			int reportsByWarehouseRow14ListCount1 = reportsRow14List.size();
			ArrayList<String> reportsByWarehouseRow14ListArray1 = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow14ListCount1;i++)
			{
				String data = reportsRow14List.get(i).getText();
				reportsByWarehouseRow14ListArray1.add(data);
			}
			String actRow14List1 = reportsByWarehouseRow14ListArray1.toString();
			String expRow14List1 = "[166, 07/09/2021, 1.0000, 1.0000, , 2.0000, 200.0000, FIFO-BIN]";
			
			int reportsByWarehouseRow15ListCount = reportsRow15List.size();
			ArrayList<String> reportsByWarehouseRow15ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow15ListCount;i++)
			{
				String data = reportsRow15List.get(i).getText();
				reportsByWarehouseRow15ListArray.add(data);
			}
			String actRow15List = reportsByWarehouseRow15ListArray.toString();
			String expRow15List = "[167, Total, 870.0000, 425.0000, 433.0000, 2.0000, 200.0000, ]";
			
			System.out.println("===========Stock Statement Report for FIFO Item HYD Location====================");
			
			System.out.println(actRow14List1);
			System.out.println(expRow14List1);
			
			System.out.println(actRow15List);
			System.out.println(expRow15List);
			
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockStatementWarehouseField2);
			Thread.sleep(1000);
			stockStatementFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockStatementFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockStatementFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(23000);
			
			int reportsByWarehouseRow17ListCount = reportsRow17List.size();
			ArrayList<String> reportsByWarehouseRow17ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow17ListCount;i++)
			{
				String data = reportsRow17List.get(i).getText();
				reportsByWarehouseRow17ListArray.add(data);
			}
			String actRow17List = reportsByWarehouseRow17ListArray.toString();
			String expRow17List = "[169, 07/09/2021, , 1.0000, , 1.0000, 200.0000, FIFO-BIN]";
			
			int reportsByWarehouseRow18ListCount = reportsRow18List.size();
			ArrayList<String> reportsByWarehouseRow18ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow18ListCount;i++)
			{
				String data = reportsRow18List.get(i).getText();
				reportsByWarehouseRow18ListArray.add(data);
			}
			String actRow18List = reportsByWarehouseRow18ListArray.toString();
			String expRow18List = "[170, Total, 1,275.0000, 425.0000, 434.0000, 1.0000, 200.0000, ]";
			
			System.out.println("===========Stock Statement Report for FIFO Item SCE Location====================");
			
			System.out.println(actRow17List);
			System.out.println(expRow17List);
			
			System.out.println(actRow18List);
			System.out.println(expRow18List);
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(3000);
			
			
			if(actRow13List.equalsIgnoreCase(expRow13List) && actRow14List.equalsIgnoreCase(expRow14List)
					&& actRow14List1.equalsIgnoreCase(expRow14List1) && actRow15List.equalsIgnoreCase(expRow15List)
					&& actRow17List.equalsIgnoreCase(expRow17List) && actRow18List.equalsIgnoreCase(expRow18List))
			{
				click(report_CloseBtn);				
				return true;
			}
			else
			{
				click(report_CloseBtn);							
				return false;
			}
			
			
			
			
			
		}
		
		
		public boolean checkStockStatementReportForRmaItem()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

			Thread.sleep(3000);

			if (fifoItemCheckBox.isSelected() == true) {
				click(fifoItemCheckBox);

			}

			if (brBatchItemCheckBox.isSelected() == true) {
				click(brBatchItemCheckBox);
			}

			if (rmaItemCheckBox.isSelected() == false) {
				click(rmaItemCheckBox);

			}
			
			dropDown(sl_DateOptionDropdown, "1");		
			Thread.sleep(2000);
			click(sl_OkBtn);
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(23000);
			
			int reportsByWarehouseRow14ListCount = reportsRow14List.size();
			ArrayList<String> reportsByWarehouseRow14ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow14ListCount;i++)
			{
				String data = reportsRow14List.get(i).getText();
				reportsByWarehouseRow14ListArray.add(data);
			}
			String actRow14List = reportsByWarehouseRow14ListArray.toString();
			String expRow14List = "[185, 07/09/2021, 1.0000, 2.0000, , 3.0000, 3,150.9528, WA-RMA]";
			
			
			int reportsByWarehouseRow15ListCount = reportsRow15List.size();
			ArrayList<String> reportsByWarehouseRow15ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow15ListCount;i++)
			{
				String data = reportsRow15List.get(i).getText();
				reportsByWarehouseRow15ListArray.add(data);
			}
			String actRow15List = reportsByWarehouseRow15ListArray.toString();
			String expRow15List = "[186, Total, 2,430.0000, 851.0000, 868.0000, 3.0000, 3,150.9528, ]";
			
			System.out.println("===========Stock Statement Report for RMA Item ====================");
			
			System.out.println(actRow14List);
			System.out.println(expRow14List);
			
			System.out.println(actRow15List);
			System.out.println(expRow15List);
			
			click(stockLedgerFilterBtn);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedgerFilterText));

			boolean actFilterpopupIsDisplaying = stockLedgerFilterText.isDisplayed();
			boolean expFilterpopupIsDisplaying = true;

			System.out.println("Filter Popup isDisplaying  : " + actFilterpopupIsDisplaying + "  Value Expected  "
					+ expFilterpopupIsDisplaying);
			
			
			click(stockStatementFilterWarehouseField);
			enterText(stockStatementFilterWarehouseField, "HYD");
			Thread.sleep(2000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(23000);
			
			
			
			
			int reportsByWarehouseRow15ListCount1 = reportsRow15List.size();
			ArrayList<String> reportsByWarehouseRow15ListArray1 = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow15ListCount1;i++)
			{
				String data = reportsRow15List.get(i).getText();
				reportsByWarehouseRow15ListArray1.add(data);
			}
			String actRow15List1 = reportsByWarehouseRow15ListArray1.toString();
			String expRow15List1 = "[167, 07/09/2021, 1.0000, 1.0000, , 2.0000, 1,375.8648, WA-RMA]";
			
			int reportsByWarehouseRow16ListCount = reportsRow16List.size();
			ArrayList<String> reportsByWarehouseRow16ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow16ListCount;i++)
			{
				String data = reportsRow16List.get(i).getText();
				reportsByWarehouseRow16ListArray.add(data);
			}
			String actRow16List = reportsByWarehouseRow16ListArray.toString();
			String expRow16List = "[168, Total, 873.0000, 426.0000, 434.0000, 2.0000, 1,375.8648, ]";
			
			
			
			System.out.println("===========Stock Statement Report for RMA Item HYD Location====================");
			
			
			
			System.out.println(actRow15List1);
			System.out.println(expRow15List1);
			
			System.out.println(actRow16List);
			System.out.println(expRow16List);
			
			
			click(report_CloseBtn);
			Thread.sleep(1000);
			click(stockLedgerHomePageFilterBtn);
			Thread.sleep(2000);
			click(stockStatementWarehouseField2);
			Thread.sleep(1000);
			stockStatementFilterWarehouseField3.sendKeys(Keys.CONTROL, "a");
			Thread.sleep(1000);
			stockStatementFilterWarehouseField3.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1000);
			enterText(stockStatementFilterWarehouseField3, "SEC");
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			Thread.sleep(1000);
			click(sl_OkBtn);
			
			Thread.sleep(2000);
			click(stockLedgerLastBtn);
			Thread.sleep(23000);
			
			int reportsByWarehouseRow17ListCount = reportsRow17List.size();
			ArrayList<String> reportsByWarehouseRow17ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow17ListCount;i++)
			{
				String data = reportsRow17List.get(i).getText();
				reportsByWarehouseRow17ListArray.add(data);
			}
			String actRow17List = reportsByWarehouseRow17ListArray.toString();
			String expRow17List = "[169, 07/09/2021, , 1.0000, , 1.0000, 1,414.2531, WA-RMA]";
			
			int reportsByWarehouseRow18ListCount = reportsRow18List.size();
			ArrayList<String> reportsByWarehouseRow18ListArray = new ArrayList<String>();
			for(int i=0;i<reportsByWarehouseRow18ListCount;i++)
			{
				String data = reportsRow18List.get(i).getText();
				reportsByWarehouseRow18ListArray.add(data);
			}
			String actRow18List = reportsByWarehouseRow18ListArray.toString();
			String expRow18List = "[170, Total, 1,275.0000, 425.0000, 434.0000, 1.0000, 1,414.2531, ]";
			
			System.out.println("===========Stock Statement Report for RMA Item SCE Location====================");
			
			System.out.println(actRow17List);
			System.out.println(expRow17List);
			
			System.out.println(actRow18List);
			System.out.println(expRow18List);
			
			click(stockLedgerFilterBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterRefreshBtn);
			Thread.sleep(1000);
			click(stockLedgerFilterOkBtn);
			
			Thread.sleep(3000);
			
			
			if(actRow14List.equalsIgnoreCase(expRow14List) && actRow15List.equalsIgnoreCase(expRow15List)
					&& actRow15List1.equalsIgnoreCase(expRow15List1) && actRow16List.equalsIgnoreCase(expRow16List)
					&& actRow17List.equalsIgnoreCase(expRow17List) && actRow18List.equalsIgnoreCase(expRow18List))
			{
				click(report_CloseBtn);				
				Thread.sleep(2000);
				click(sl_CloseBtn);				
				return true;
			}
			else
			{
				click(report_CloseBtn);				
				Thread.sleep(2000);
				click(sl_CloseBtn);							
				return false;
			}
			
			
			
			
			
		}
		//////////////////////////////////
		//////////////////////////
		// Stock In Stock Out Vouchers

	public boolean checkSavingStockOutVoucher()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		
		Thread.sleep(3000);
		click(userNameDisplay);
		Thread.sleep(1000);
		click(logoutOption);		
		Thread.sleep(3000);
		
		String unamelt="su";

		String pawslt="su";
		
		LoginPage.enterUserName(unamelt);	     
        LoginPage.enterPassword(pawslt);
        
		
		 Select company = new Select(companyDropDownList);
		 company.selectByValue("108");
		 Thread.sleep(2000);
		 
  
        LoginPage.clickOnSignInBtn();  
        
        Thread.sleep(3000);
        
        System.out.println("Signing");
        
        // Login User Name is Verify
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		String actUserInfo                      = userNameDisplay.getText();
		String expUserInfo                      = "SU";	

		System.out.println("Login With Valid User and Check User Name Display Value Actual : ." + actUserInfo + ". Value Expected : " + expUserInfo);
		
		
		// Login Company Name is Verify
		Thread.sleep(1000);
		click(companyLogo);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
		String actGetLoginCompanyNameInformation = companyName.getText();
		String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
		String expGetLoginCompanyName            = "Maintain Quantity a";   
		companyLogo.click();
		
		System.out.println("Company Name Display Value Actual                : ." + actGetLoginCompanyName +  ". Value Expected : " + expGetLoginCompanyName);
      
		Thread.sleep(2000);

		click(inventoryMenu);
		Thread.sleep(2000);
		click(inventoryTransactionsMenu);
		click(inventoryTransactionsStockMenu);
		click(inventoryTransactionsStockOut);
		click(newBtn);
		
		checkValidationMessage("Screen opened");

		enterText(stockAccounttxt, "Customer A");
		Thread.sleep(2000);
		stockAccounttxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		enterText(partyAccounttxt, "Sales - Computers");
		Thread.sleep(2000);
		partyAccounttxt.sendKeys(Keys.TAB);
		enterText(Warehouse1txt, "HYD");
		Thread.sleep(2000);
		Warehouse1txt.sendKeys(Keys.TAB);
		enterText(Warehouse2txt, "SEC");
		Thread.sleep(2000);
		Warehouse2txt.sendKeys(Keys.TAB);
		enterText(stockOutDepartmenttxt, "DEPT1");
		Thread.sleep(2000);
		stockOutDepartmenttxt.sendKeys(Keys.TAB);
		click(select1stRow_1stColumn);
		enterText(enterItemttxt, "Stock");
		Thread.sleep(2000);
		enterItemttxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		enterText(enterQuantitytxt, "1");
		Thread.sleep(2000);
		enterQuantitytxt.sendKeys(Keys.TAB);
		Thread.sleep(1000); 
		if (validationConfirmationMessage.getText().isEmpty()==false) 
		{
			checkValidationMessage("This Transaction will make the Stock Negative");
		}
		Thread.sleep(1000);
		enterRatetxt.clear();
		Thread.sleep(1000);
		enterText(enterRatetxt, "100");
		Thread.sleep(2000);
		enterRatetxt.sendKeys(Keys.TAB);
		
		
		Thread.sleep(1000);
		click(voucherSaveBtn);

		Thread.sleep(3000);
				
		
		HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 4; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("This Transaction will make the Stock Negative");
		expMessage.add("Saving in background.");
		expMessage.add("This Transaction will make the Stock Negative");
		expMessage.add("Voucher saved successfully : 1");

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if (actMessage.equals(expMessage)) {

			
			return true;
		} else {
			
			return false;
		}
	}

	
	public boolean checkSavingStockInVoucher()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		Thread.sleep(1500);
		
		documentNumberTxt.click();
		Thread.sleep(1000);
		documentNumberTxt.sendKeys(Keys.CONTROL,"a");
		Thread.sleep(1000);		
		documentNumberTxt.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(1000);
		documentNumberTxt.sendKeys("2");
		Thread.sleep(1000);
		documentNumberTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		click(previousBtn);
		checkValidationMessage("Voucher loaded successfully");
		
		Thread.sleep(1000);
		click(toggleBtn);
		Thread.sleep(1000);
		
		moveToElement(copyToClipBoardOption);
		Thread.sleep(1000);
		
		enterText(homepageManuSearchField, "Stock in");
		Thread.sleep(1000);	
		moveToElement(stockIn);
		//Thread.sleep(1000);	
		click(newBtn);
		//Thread.sleep(1000);	
		checkValidationMessage("Screen opened");
		//Thread.sleep(1000);	
		click(toggleBtn);
		Thread.sleep(1000);
		click(pasteFromClipBoardOption);
		//Thread.sleep(2000);
		
		checkValidationMessage("Paste from clipboard completed successfully");
		Thread.sleep(1000);
		enterText(Warehouse2txt, "SEC");
		Thread.sleep(1000);
		tab(Warehouse2txt);
		/*
		click(select1stRow_7thColumn);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));

		boolean actBatchpopupIsDisplaying = batchPickOnFIFOIcon.isDisplayed();
		boolean expBatchpopupIsDisplaying = true;

		System.out.println("Batch Popup isDisplaying  : " + actBatchpopupIsDisplaying + "  Value Expected  "
				+ expBatchpopupIsDisplaying);
		click(batchPickOnFIFOIcon);
		click(batchOkIcon); */

		Thread.sleep(1000);
		click(voucherSaveBtn);
		/*
		Thread.sleep(2000);

		HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) {
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage); */
		
		
		

		//Thread.sleep(1000);
		
		
		HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 4; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("This Transaction will make the Stock Negative");
		expMessage.add("Saving in background.");
		expMessage.add("This Transaction will make the Stock Negative");
		expMessage.add("Voucher saved successfully : 1");

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);


		if (actMessage.equals(expMessage)) {

			click(new_CloseBtn);
			Thread.sleep(1000);
			click(voucherhomeCloseBtn);
			return true;
		} else {
			click(new_CloseBtn);
			Thread.sleep(1000);
			click(voucherhomeCloseBtn);
			return false;
		}
	}
	
	
	
	public boolean checkStockLedgerReportForStockItemAfterStockInAndStockOutVouchers()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryStockLedgerMenu);
		Thread.sleep(3000);

		int rowcount = stockLedgerHometableRowCount.size();

		System.out.println(rowcount);

		for (int i = 0; i < rowcount; i++) {
			String actName = ledgerHometableItemNamesList.get(i).getText();

			System.out.println(actName);

			if (actName.equalsIgnoreCase("Stock")) {
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		dropDown(sl_DateOptionDropdown, "6");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		
		
		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = "[2, "+getCurrentDateF2()+", Stock inA:1, HYD, , , 1.0000, , 1.0000, , , , , , , ]";
		
		
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		int reportsByWarehouseRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsByWarehouseRow3ListArray.add(data);
		}
		String actRow3List = reportsByWarehouseRow3ListArray.toString();
		String expRow3List = "[3, "+getCurrentDateF2()+", Stock inA:1, SEC, 1.0000, , , , , , , , , , , ]";
		
		
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		
		
		

		if(actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);				
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);							
			return false;
		}
	}
	
	
	
	public boolean checkLoginToTest1UserAndAuthorizeStockOutVocher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(3000);
		click(userNameDisplay);
		Thread.sleep(1000);
		click(logoutOption);		
		Thread.sleep(3000);
		
		String unamelt="Test1";

		String pawslt="su";
		
		LoginPage.enterUserName(unamelt);	     
        LoginPage.enterPassword(pawslt);
        
		
		 Select company = new Select(companyDropDownList);
		 company.selectByValue("108");
		 Thread.sleep(2000);
		 
  
        LoginPage.clickOnSignInBtn();  
        
        Thread.sleep(3000);
        
        System.out.println("Signing");
        
        // Login User Name is Verify
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		String actUserInfo                      = userNameDisplay.getText();
		String expUserInfo                      = "TEST1";	

		System.out.println("Login With Valid User and Check User Name Display Value Actual : ." + actUserInfo + ". Value Expected : " + expUserInfo);
		
		
		// Login Company Name is Verify
		Thread.sleep(1000);
		click(companyLogo);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
		String actGetLoginCompanyNameInformation = companyName.getText();
		String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
		String expGetLoginCompanyName            = "Maintain Quantity a";   
		companyLogo.click();
		
		System.out.println("Company Name Display Value Actual                : ." + actGetLoginCompanyName +  ". Value Expected : " + expGetLoginCompanyName);
      
        
        
        click(inventoryMenu);
		click(inventoryTransactionsMenu);
		click(inventoryTransactionsStockMenu);
		click(inventoryTransactionsStockOut);
		Thread.sleep(2000);
		click(stockOutPendingAuthLink);
		
		
		Thread.sleep(3000);
		
		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase("1")) 
			{
				grid_CheckBoxList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		click(stockOutHomeAuthBtn);
		
		Thread.sleep(2000);
		
		click(homePageAllVouchersView);
		
		Thread.sleep(2000);
		
		String actRow1AuthorizationStatus = voucherGridRow1AuthorizationStatus.getText();
		String expRow1AuthorizationStatus = "Authorized";
		
		
		System.out.println("Row1AuthorizationStatus  : "+actRow1AuthorizationStatus  +" Value Expected : "+expRow1AuthorizationStatus);
		
		
		if(actRow1AuthorizationStatus.equalsIgnoreCase(expRow1AuthorizationStatus))
		{
			click(voucherhomeCloseBtn);			
			Thread.sleep(1000);
			return true;
		}
		else
		{
			click(voucherhomeCloseBtn);
			Thread.sleep(1000);
			return false;
		}
	}
	
	
	public boolean checkStockLedgerReportForStockItemAfterAutrhroizedStockOutVoucher()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryStockLedgerMenu);
		Thread.sleep(3000);

		int rowcount = stockLedgerHometableRowCount.size();

		System.out.println(rowcount);

		for (int i = 0; i < rowcount; i++) {
			String actName = ledgerHometableItemNamesList.get(i).getText();

			System.out.println(actName);

			if (actName.equalsIgnoreCase("Stock")) {
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		dropDown(sl_DateOptionDropdown, "6");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		
		
		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = "[2, "+getCurrentDateF2()+", Stock outA:1, HYD, , , 1.0000, , 1.0000, , , , , , , ]";
		
		
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		int reportsByWarehouseRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsByWarehouseRow3ListArray.add(data);
		}
		String actRow3List = reportsByWarehouseRow3ListArray.toString();
		String expRow3List = "[3, "+getCurrentDateF2()+", Stock outA:1, SEC, 1.0000, , , , , , , , , , , ]";
		
		
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		int reportsByWarehouseRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsByWarehouseRow4ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsByWarehouseRow4ListArray.add(data);
		}
		String actRow4List = reportsByWarehouseRow4ListArray.toString();
		String expRow4List = "[4, "+getCurrentDateF2()+", Stock inA:1, HYD, , , 1.0000, , 1.0000, , , , , , , ]";
		
		
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		int reportsByWarehouseRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsByWarehouseRow5ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsByWarehouseRow5ListArray.add(data);
		}
		String actRow5List = reportsByWarehouseRow5ListArray.toString();
		String expRow5List = "[5, "+getCurrentDateF2()+", Stock inA:1, SEC, 1.0000, , , , , , , , , , , ]";
		
		
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		
		
		

		if(actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);				
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);							
			return false;
		}
	}
	
	
	
	public boolean checkStockLedgerReportAfterSavingSalesReturnsVocherWithRefineFilterFirstPage()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	
	{
		
		
		reLogin("su", "su");
		
		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryStockLedgerMenu);
		Thread.sleep(3000);

	
		click(selectAllItemsCheckBox);
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		/*click(stockLedgerLastBtn);
		Thread.sleep(10000);*/
		
		
		FilterRefine("Voucher", "PurVou:1");
		
		Thread.sleep(3000);
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1, , Opening Balance, , 20.0000, , , , 20.0000, , , 202.1210, 10.1061, , , ]";
		
		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = "[2, 01/10/2020, PurVou:1, HYD, 10.0000, 20.2020, , , 30.0000, 202.0200, , 404.1410, 13.4714, 01102020PH, , ]";
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = "[3, 01/10/2020, PurVou:1, SEC, 10.0000, 30.3030, , , 40.0000, 303.0300, , 707.1710, 17.6793, 01102020PS, , ]";
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = "[4, 01/10/2020, PurVou:1, TEST, 10.0000, 40.4040, , , 50.0000, 404.0400, , 1,111.2110, 22.2242, 01102020PT, , ]";
		
		
		
		System.out.println("*********************************checkStockLedgerReport Using Refine Filter 1st Page Report*****************************************");
		
	
		
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		
		
	
		
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
				&& actRow4List.equalsIgnoreCase(expRow4List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);				
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	
	public boolean checkStockLedgerReportAfterSavingSalesReturnsVocherWithRefineFilterLastPage()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	
	{
		
		
	
		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryStockLedgerMenu);
		Thread.sleep(3000);

	
		click(selectAllItemsCheckBox);
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(8000);
		
		
		FilterRefine("Voucher", "PurVou:35");
		
		Thread.sleep(3000);
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1920, , Opening Balance, , 20.0000, , , , 20.0000, , , 202.1210, 10.1061, , , ]";
		
		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = "[1921, 01/09/2021, PurVou:35, HYD, 10.0000, 1,000.0000, , , 15.0000, 10,000.0000, , 10,040.3131, 669.3542, 01092021PH, , ]";
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = "[1922, 01/09/2021, PurVou:35, SEC, 10.0000, 2,000.0000, , , 25.0000, 20,000.0000, , 30,040.3131, 1,201.6125, 01092021PH, , ]";
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = "[1923, , Opening Balance, , 20.0000, , , , 20.0000, , , 202.1210, 10.1061, , , ]";
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = "[1924, 01/09/2021, PurVou:35, HYD, 10.0000, 1,000.0000, , , 15.0000, 10,000.0000, , 10,041.0807, 669.4054, , BIN1-10.00, ]";
		
		
		int report6thRowListCount = reportsRow6List.size();
		ArrayList<String> report6thRowListArray = new ArrayList<String>();
		for(int i=0;i<report6thRowListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			report6thRowListArray.add(data);
		}
		String actRow6List = report6thRowListArray.toString();
		String expRow6List = "[1925, 01/09/2021, PurVou:35, SEC, 10.0000, 2,000.0000, , , 25.0000, 20,000.0000, , 30,041.0807, 1,201.6432, , BIN2-10.00, ]";
		
		
		int report7thRowListCount = reportsRow7List.size();
		ArrayList<String> report7thRowListArray = new ArrayList<String>();
		for(int i=0;i<report7thRowListCount;i++)
		{
			String data = reportsRow7List.get(i).getText();
			report7thRowListArray.add(data);
		}
		String actRow7List = report7thRowListArray.toString();
		String expRow7List = "[1926, , Opening Balance, , 20.0000, , , , 20.0000, , , 202.1210, 10.1061, , , ]";
	
		
	
		
		int report8thRowListCount = reportsRow8List.size();
		ArrayList<String> report8thRowListArray = new ArrayList<String>();
		for(int i=0;i<report8thRowListCount;i++)
		{
			String data = reportsRow8List.get(i).getText();
			report8thRowListArray.add(data);
		}
		String actRow8List = report8thRowListArray.toString();
		String expRow8List = "[1927, 01/09/2021, PurVou:35, HYD, 10.0000, 1,000.0000, , , 15.0000, 10,000.0000, , 10,048.0097, 669.8673, , , 01092021PS, 01092021PS1, 01092021PS2, 01092021PS3, 01092021PS4, 01092021PS5, 01092021PS6, 01092021PS7, 01092021PS8, 01092021PS9]";
		
		
		
		int report9thRowListCount = reportsRow9List.size();
		ArrayList<String> report9thRowListArray = new ArrayList<String>();
		for(int i=0;i<report9thRowListCount;i++)
		{
			String data = reportsRow9List.get(i).getText();
			report9thRowListArray.add(data);
		}
		String actRow9List = report9thRowListArray.toString();
		String expRow9List = "[1928, 01/09/2021, PurVou:35, SEC, 10.0000, 2,000.0000, , , 25.0000, 20,000.0000, , 30,048.0097, 1,201.9204, , , 01092021PH10, 01092021PH11, 01092021PH12, 01092021PH13, 01092021PH14, 01092021PH15, 01092021PH16, 01092021PH17, 01092021PH18, 01092021PH9]";
		
		
		
		/*int report10thRowListCount = reportsRow10List.size();
		ArrayList<String> report10thRowListArray = new ArrayList<String>();
		for(int i=0;i<report10thRowListCount;i++)
		{
			String data = reportsRow10List.get(i).getText();
			report10thRowListArray.add(data);
		}
		String actRow10List = report10thRowListArray.toString();
		String expRow10List = "";
		
		
		
		int reportsRow11ListCount = reportsRow11List.size();
		ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow11ListCount;i++)
		{
			String data = reportsRow11List.get(i).getText();
			reportsRow11ListArray.add(data);
		}
		String actRow11List = reportsRow11ListArray.toString();
		String expRow11List = "";*/
		
		
		
		System.out.println("*********************************checkStockLedgerReport Using Refine Filter 2nd Page Report*****************************************");
		
	
		
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		System.out.println(actRow7List);
		System.out.println(expRow7List);
		
		System.out.println(actRow8List);
		System.out.println(expRow8List);
		
		System.out.println(actRow9List);
		System.out.println(expRow9List);
		
		/*System.out.println(actRow10List);
		System.out.println(expRow10List);
		
		System.out.println(actRow11List);
		System.out.println(expRow11List);*/
		
		
		
		
		
		
		
		
		
		
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List)
				&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
				/*&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List)*/)
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);				
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	public void FilterRefine(String str1, String str2) throws InterruptedException{
		
		click(stockLedgerFilterExpandBtn);
		Thread.sleep(1500);
		click(stockLedgerFilterRefineBtn);
		//Thread.sleep(18000);
		
		new WebDriverWait(getDriver(), 50).until(ExpectedConditions.elementToBeClickable(filterRefineLookInDropdownBtn));
		dropDown(str1, filterRefineLookInDropdownBtn);
		
		Thread.sleep(2000);
		
		//dropDown(str2, filterRefineLookIn2DropdownBtn);
		
		enterText(filterRefineFindWhatTxt, str2);
		
		
		Thread.sleep(2000);
		
		click(filterRefineOkBtn);
		
		
		
	}
	
	
	
	public boolean checkStockMovementReportWithRefineFilter()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		
		
		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryReportsStockMovement);
		//click(sl_HeaderSelectChkBox);
		click(selectAllItemsCheckBox);
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		
		
		FilterRefine("Item", "FIFO-BIN");
		
		
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1, FIFO-BIN, 20.00, 202.1210, 1,076.00, 51,459.1768, 1,091.00, -49,061.2978, 5.00, 2,600.0000, 520.0000]";
		
		

		
		
		
	
		
		System.out.println("*********************************checkStockMovementReport With Filter Refine*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		
		
	
		
		if(actRow1List.equalsIgnoreCase(expRow1List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	
	public boolean checkStockValuationReportByProductWithRefineFilter()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryReportsStockValuation);
		//click(sl_HeaderSelectChkBox);
		click(selectAllItemsCheckBox);
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		
		
		
		FilterRefine("Item", "WA-RMA");
		
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1, WA-RMA, WA-RMA, WA-RMA, 5.0000, 5,507.9347, 1,101.5869, , ]";
		
		
		
		System.out.println("*********************************checkStockValuationReport With Filter refine*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		
		
	
		
		if(actRow1List.equalsIgnoreCase(expRow1List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	
	public boolean checkStockAgeingAnalysisReportByDefaultWithRefineFilter()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		click(inventoryMenu);
		click(InventoryReportsMenu);
		click(InventoryReportsStockAgeingAnalysisMenu);
		click(InventoryReportsAgeingAnalysis);
		click(sl_HeaderSelectChkBox);
		//click(selectAllItemsCheckBox);
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		
		
		
		FilterRefine("Particulars", "BR-BATCH");
		
		Thread.sleep(3000);
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<4;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1, BR-BATCH, 5.00, 4,150.5050]";
		
	
		
		System.out.println("*********************************checkStockAgeing AnalysisReport With Filter Refine*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		
		
	
		
		if(actRow1List.equalsIgnoreCase(expRow1List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	
	public boolean checkStockBalanceByWarehouseReportWithRefineFilter()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		
		click(inventoryMenu);
		click(InventoryReportsMenu);
		Thread.sleep(2000);
		moveToElement(InventoryReportsStockBalanceByWarehouseMenu);
		click(sl_HeaderSelectChkBox);
		//click(selectAllItemsCheckBox);
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		
		
		FilterRefine("Quantity", "2.0000000");
		
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1, FIFO-BIN, 0, 0, 2.0000000, , 2.0000000, 100.00, 520.0000]";
		
		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = "[2, BR-BATCH, 0, 0, 2.0000000, , 2.0000000, 60.10, 830.1010]";
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = "[3, WA-RMA, 0, 0, 2.0000000, , 2.0000000, 687.93, 1,101.5869]";
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = "[4, WA-RMA, 0, 0, 2.0000000, , 2.0000000, 687.93, 1,101.5869]";
		
	
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = "[5, FIFO-BIN, 0, 0, 2.0000000, , 2.0000000, 100.00, 520.0000]";
		
		
		int report6thRowListCount = reportsRow6List.size();
		ArrayList<String> report6thRowListArray = new ArrayList<String>();
		for(int i=0;i<report6thRowListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			report6thRowListArray.add(data);
		}
		String actRow6List = report6thRowListArray.toString();
		String expRow6List = "[6, BR-BATCH, 0, 0, 2.0000000, , 2.0000000, 60.10, 830.1010]";
		
		
		
		
		
		
		System.out.println("*********************************checkStockBalanceByWarehouseReport*****************************************");
		
	
		
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		
		
		
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) &&
				actRow6List.equalsIgnoreCase(expRow6List))
		{
					
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);		
			return false;
		}
    
	
	
	}
	
	
	public boolean checkLedgerReportWithRefineFilterFirstPage()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	
	{
		
		click(financialsMenu);
		click(financialReports);
		click(financialReportsLedger);
		Thread.sleep(3000);

	
		click(selectAllItemsCheckBox);
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
	
		
		FilterRefine("Debit", "1,500.0000000");
		
		Thread.sleep(6200);
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1, 02/09/2021, SalInv : 35, Sales - Computers, 1,500.0000000, , 86,808.7907000, , , , 1,500.0000000, , 86,808.7907000, Indian Rupees, BR-BATCH]";
		
		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = "[2, 02/09/2021, SalInv : 35, Sales - Computers, 1,500.0000000, , 85,308.7907000, , , , 1,500.0000000, , 85,308.7907000, Indian Rupees, BR-BATCH]";
		
		
		
		
		
		
		System.out.println("*********************************check LedgerReport Using Refine Filter 1st Page Report*****************************************");
		
	
		
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);				
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	
	
	public boolean checkLedgerReportWithRefineFilterLastPage()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	
	{
		
		click(financialsMenu);
		click(financialReports);
		click(financialReportsLedger);
		Thread.sleep(3000);

	
		click(selectAllItemsCheckBox);
		
		dropDown(sl_DateOptionDropdown, "1");		
		Thread.sleep(2000);
		click(sl_OkBtn);
		Thread.sleep(2000);
		click(stockLedgerLastBtn);
		Thread.sleep(6000);
	
		
		FilterRefine("Balance", "3,310.0000000");
		
		Thread.sleep(6000);
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[552, 07/11/2020, SalInv : 5, Sales - Computers, 825.0000000, , 3,310.0000000, , , , 825.0000000, , 3,310.0000000, Indian Rupees, BR-BATCH]";
		
		

		
		
		
		System.out.println("*********************************check LedgerReport Using Refine Filter Last Page Report*****************************************");
		
	
		
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		
	
		
		if(actRow1List.equalsIgnoreCase(expRow1List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);				
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	public boolean checkPurchaseRegisterReportWithRefineFilter()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	
	{

		click(financialsMenu);
		click(financialReports);
		click(financialPurchaseReportMenu);
		click(financialPurchaseRegister);
		Thread.sleep(3000);

		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(2000);

		//dropDown("Cost of goods sold - Computers", accountDropdown);
		enterText(accountDropdown, "Cost of goods sold - Computers");
		Thread.sleep(2000);

		click(sl_OkBtn);
		Thread.sleep(2000);

	
		FilterRefine("Party", "Bank");

		Thread.sleep(3000);
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1, 01/09/2021, PurVou : 35, Bank, 10,000.00, BR-BATCH, 10.00, 1,000.00]";
		
		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = "[2, 01/09/2021, PurVou : 35, Bank, 20,000.00, BR-BATCH, 10.00, 2,000.00]";
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = "[3, 01/09/2021, PurVou : 35, Bank, 10,000.00, FIFO-BIN, 10.00, 1,000.00]";
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = "[4, 01/09/2021, PurVou : 35, Bank, 20,000.00, FIFO-BIN, 10.00, 2,000.00]";
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=0;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = "[5, 01/09/2021, PurVou : 35, Bank, 10,000.00, WA-RMA, 10.00, 1,000.00]";
		
		
		int report6thRowListCount = reportsRow6List.size();
		ArrayList<String> report6thRowListArray = new ArrayList<String>();
		for(int i=0;i<report6thRowListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			report6thRowListArray.add(data);
		}
		String actRow6List = report6thRowListArray.toString();
		String expRow6List = "[6, 01/09/2021, PurVou : 35, Bank, 20,000.00, WA-RMA, 10.00, 2,000.00]";

		
		
		
		System.out.println("*********************************check Purchase Register Report Using Refine Filter Last Page Report*****************************************");
		
	
		
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
	
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);				
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	public boolean checkTrailBalanceReportWithRefineFilter()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	
	{

		click(financialsMenu);
		click(financialFinalAccountsMenu);
		click(trialBalance);
		
		Thread.sleep(1500);

		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(1000);

		click(sl_OkBtn);
		Thread.sleep(2000);

	
		FilterRefine("Alias", "Cash & bank");

		Thread.sleep(3000);
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1, Cash & bank, 121, Cash & bank, , 82,082.97, , , , 82,082.97, 9,383.27, 91,466.24, , 82,082.97, , , , 82,082.97, 9,383.27, 91,466.24, , , , , , , , , 76.88, 76.88, ]";
		
		

		

		
		
		
		System.out.println("*********************************check Trial Report Using Refine Filter*****************************************");
		
	
		
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		
		
	
		
		if(actRow1List.equalsIgnoreCase(expRow1List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);				
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	public boolean checkCustomerAgeingSummeryReportWithRefineFilter()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	
	{

		click(financialsMenu);
		click(financialReceivableAndPayableMenu);
		click(customerSummeryMenu);
		click(customerAgeingSummery);
		Thread.sleep(1500);
		
		click(selectAllItemsCheckBox);
		
		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(1000);

		click(sl_OkBtn);
		Thread.sleep(2000);
		
		

	
		FilterRefine("Account", "Rathod BN");

		Thread.sleep(3000);
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			
			if(data.isEmpty()==false)
			{
				reportsByWarehouseRow1ListArray.add(data);
			}
			
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1, Rathod BN, 1,405.45, 1,405.45, 1,405.45, 1,405.45, 1,405.45, 1,405.45, 1,405.45, 1,405.45, 1,405.45, 1,405.45, 1,405.45, 1,405.45, RBN]";
		
							  
		
		System.out.println("*********************************check Customer Ageing Summary Report Using Refine Filter*****************************************");
		
	
		
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		
		
	
		
		if(actRow1List.equalsIgnoreCase(expRow1List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);				
			return true;
		}
		else
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	public boolean checkVendorAgeingSummeryReportWithRefineFilter()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	
	{

		click(financialsMenu);
		click(financialReceivableAndPayableMenu);
		click(vendorSummeryMenu);
		click(vendorAgeingSummery);
		Thread.sleep(1500);
		
		click(selectAllItemsCheckBox);
		
		dropDown(sl_DateOptionDropdown, "1");
		Thread.sleep(1000);

		click(sl_OkBtn);
		Thread.sleep(2000);
		
		

	
		FilterRefine("Balance Amount", "2,400.09");

		Thread.sleep(3000);
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			
			if(data.isEmpty()==false)
			{
				reportsByWarehouseRow1ListArray.add(data);
			}
			
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = "[1, DHANA Bnk, 2,400.09, 2,400.09, 2,400.09, 2,400.09, 2,400.09, 2,400.09, 2,400.09, 2,400.09, 2,400.09, 2,400.09, 2,400.09, 2,400.09, DHANA Bnk]";
		
		

		

		
		
		
		System.out.println("*********************************check Customer Ageing Summary Report Using Refine Filter*****************************************");
		
	
		
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		
		
	
		
		if(actRow1List.equalsIgnoreCase(expRow1List))
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);				
			return true;
		}
		
		{
			click(report_CloseBtn);				
			Thread.sleep(2000);
			click(sl_CloseBtn);
			return false;
		}
	}
	
	
	
	
	

		

		

	}
		
		
		
		
	
		
	
	
	




	 




					    
					    
						
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
		 
