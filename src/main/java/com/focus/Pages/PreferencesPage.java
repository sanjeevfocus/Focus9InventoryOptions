package com.focus.Pages;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.JavascriptExecutor;
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

public class PreferencesPage extends BaseEngine {

	public PreferencesPage(WebDriver driver) {
		PageFactory.initElements(driver, this);

	}
	static String actFileName;


	// Preferences TDS Elements

	@FindBy(xpath = "//div[@id='btnSubmit']")
	private static WebElement preferences_SaveBtn;

	@FindBy(xpath = "(//span[@class='icon-close data-accept'])[1]")
	private static WebElement preferences_CloseBtn;

	@FindBy(xpath = "//button[@id='TDSdefaultOpen']")
	private static WebElement preferences_TDSTab;

	@FindBy(xpath = "//input[@id='chkTDSbyTags']")
	private static WebElement preferences_TdsByTagsCheckBox;

	@FindBy(xpath = "//select[@id='TDSByTagsdropdwonId']")
	private static WebElement preferences_TdsByTagsDropDown;

	@FindBy(xpath = "//input[@id='rdoNonPublicSector']")
	private static WebElement preferences_NonPublicSectorCheckBox;

	@FindBy(xpath = "//input[@id='rdoPublicSector']")
	private static WebElement preferences_PublicSectorCheckBox;

	@FindBy(xpath = "//input[@id='chkTDSfromDirectExpenses']")
	private static WebElement preferences_TDSOnDirectExpensesCheckBox;

	@FindBy(xpath = "//input[@id='rboByVendors']")
	private static WebElement preferences_ByVendorsCheckBox;

	@FindBy(xpath = "//input[@id='rboAllVendors']")
	private static WebElement preferences_AllVendorsCheckBox;

	@FindBy(xpath = "//button[@class='multiselect dropdown-toggle btn btn-default']")
	private static WebElement preferences_AllDropDown;

	@FindBy(xpath = "(//span[@class='icon-add data-accept'])[1]")
	private static WebElement preferences_VocherType;
	
	@FindBy(xpath = "//button[@id='btnTCSVoucherTypeview']")
	private static WebElement preferences_TCSVocherType;

	// Vocher Types Elements

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[1]/td[3]/input")  
	private static WebElement preferences_VocherTypeRow1CheckBox;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[2]/td[3]/input")
	private static WebElement preferences_VocherTypeRow2CheckBox;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[3]/td[3]/input")
	private static WebElement preferences_VocherTypeRow3CheckBox;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[8]/td[3]/input")
	private static WebElement preferences_VocherTypeRow8CheckBox;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[9]/td[3]/input")
	private static WebElement preferences_VocherTypeRow9CheckBox;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[4]/td[3]/input")
	private static WebElement preferences_VocherTypeRow4CheckBox;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[15/td[3]/input")
	private static WebElement preferences_VocherTypeRow5CheckBox;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[6]/td[3]/input")
	private static WebElement preferences_VocherTypeRow6CheckBox;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[7]/td[3]/input")
	private static WebElement preferences_VocherTypeRow7CheckBox;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[1]/td[5]/button")
	private static WebElement preferences_VocherTypeRow1FiedsButton;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[2]/td[5]/button")
	private static WebElement preferences_VocherTypeRow2FiedsButton;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[3]/td[5]/button")
	private static WebElement preferences_VocherTypeRow3FiedsButton;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[4]/td[5]/button")
	private static WebElement preferences_VocherTypeRow4FiedsButton;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[5]/td[5]/button")
	private static WebElement preferences_VocherTypeRow5FiedsButton;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[6]/td[5]/button")
	private static WebElement preferences_VocherTypeRow6FiedsButton;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[7]/td[5]/button")
	private static WebElement preferences_VocherTypeRow7FiedsButton;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[1]/td[4]")
	private static WebElement preferences_VocherTypeRow1VocherName;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[2]/td[4]")
	private static WebElement preferences_VocherTypeRow2VocherName;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[3]/td[4]")
	private static WebElement preferences_VocherTypeRow3VocherName;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[4]/td[4]")
	private static WebElement preferences_VocherTypeRow4VocherName;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[5]/td[4]")
	private static WebElement preferences_VocherTypeRow5VocherName;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[6]/td[4]")
	private static WebElement preferences_VocherTypeRow6VocherName;

	@FindBy(xpath = "//*[@id='saveTDSVoucher']/tr[7]/td[4]")
	private static WebElement preferences_VocherTypeRow7VocherName;

	@FindBy(xpath = "//div[@id='btnVouchertypeOK']")
	private static WebElement preferences_TDSVocherTypeOkBtn;

	@FindBy(xpath = "//div[@id='btnvouchertypeviewclose']")
	private static WebElement preferences_TDSVocherTypeCloseBtn;
	
	//TCS Vocher Elements
	
	@FindBy(xpath = "//*[@id='saveTCSVoucher']/tr[1]/td[3]/input")  
	private static WebElement preferences_TcsVocherTypeRow1CheckBox;
	
	@FindBy(xpath = "//input[@id='TCStaxselectionFields1']")  
	private static WebElement preferences_TcsVocherTypeCGSTCheckBox;
	
	@FindBy(xpath = "//input[@id='TCStaxselectionFields2']")  
	private static WebElement preferences_TcsVocherTypeSGSTCheckBox;
	

	@FindBy(xpath = "//*[@id='saveTCSVoucher']/tr[4]/td[3]/input")  
	private static WebElement preferences_TcsVocherTypeRow4CheckBox;
	

	@FindBy(xpath = "//*[@id='saveTCSVoucher']/tr[5]/td[3]/input")  
	private static WebElement preferences_TcsVocherTypeRow5CheckBox;
	
	@FindBy(xpath = "//input[@id='TCStaxselectionFields5']")  
	private static WebElement preferences_TcsVocherTaxableCheckBox;
	
	
	@FindBy(xpath = "//*[@id='saveTCSVoucher']/tr[1]/td[5]/button")  
	private static WebElement preferences_TcsVocherTypeRow1FieldsBtn;
	
	@FindBy(xpath="//*[@id='TCStaxexcemptionbody']/tr[1]/td[4]/label")
	private static WebElement  Row1editScreenAndFooterFieldsName;
	
	@FindBy(xpath="//*[@id='TCStaxexcemptionbody']/tr[2]/td[4]/label")
	private static WebElement  Row2editScreenAndFooterFieldsName;
	
	@FindBy(xpath="//*[@id='TCStaxexcemptionbody']/tr[3]/td[4]/label")
	private static WebElement  Row3editScreenAndFooterFieldsName;
	
	@FindBy(xpath="//*[@id='TCStaxexcemptionbody']/tr[4]/td[4]/label")
	private static WebElement  Row4editScreenAndFooterFieldsName;
	
	@FindBy(xpath="//*[@id='TCStaxexcemptionbody']/tr[5]/td[4]/label")
	private static WebElement  Row5editScreenAndFooterFieldsName;
	
	@FindBy(xpath="//*[@id='TCStaxexcemptionbody']/tr[6]/td[4]/label")
	private static WebElement  Row6editScreenAndFooterFieldsName;
	
	@FindBy(xpath = "//*[@id='btnTCStaxexcemptionOK']")  
	private static WebElement editScreenAndFooterOkBtn;
	
	@FindBy(xpath = "//div[@id='btnTCSVouchertypeOK']")  
	private static WebElement tcsVocherTypeOkBtn;

	// Edit Screen and Footer fields

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[1]/td[3]/input")
	private static WebElement preferences_EditAndFooterFieldrow1CheckBox;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[2]/td[3]/input")
	private static WebElement preferences_EditAndFooterFieldrow2CheckBox;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[3]/td[3]/input")
	private static WebElement preferences_EditAndFooterFieldrow3CheckBox;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[4]/td[3]/input")
	private static WebElement preferences_EditAndFooterFieldrow4CheckBox;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[5]/td[3]/input")
	private static WebElement preferences_EditAndFooterFieldrow5CheckBox;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[6]/td[3]/input")
	private static WebElement preferences_EditAndFooterFieldrow6CheckBox;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[1]/td[4]/label")
	private static WebElement preferences_EditAndFooterrow1FieldName;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[2]/td[4]/label")
	private static WebElement preferences_EditAndFooterrow2FieldName;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[3]/td[4]/label")
	private static WebElement preferences_EditAndFooterrow3FieldName;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[4]/td[4]/label")
	private static WebElement preferences_EditAndFooterrow4FieldName;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[5]/td[4]/label")
	private static WebElement preferences_EditAndFooterrow5FieldName;

	@FindBy(xpath = "//tbody[@id='savetaxexcemptionbody']/tr[6]/td[4]/label")
	private static WebElement preferences_EditAndFooterrow6FieldName;

	@FindBy(xpath = "//div[@id='btntaxexcemptionOK']")
	private static WebElement preferences_EditAndFooterFieldOkBtn;

	@FindBy(xpath = "//div[@id='btntaxexemptionclose']")
	private static WebElement preferences_EditAndFooterFieldCloseBtn;

	@FindBy(xpath = "//input[@id='chkAskforTDSDeduction']")
	private static WebElement preferences_AskForTdsDeductionCheckBox;

	@FindBy(xpath = "//input[@id='chkDefineSlab']")
	private static WebElement preferences_DefineSlabCheckBox;

	@FindBy(xpath = "//input[@id='chkmultiTDSDeduction']")
	private static WebElement preferences_AllowMultipleTdsCheckBox;

	@FindBy(xpath = "//select[@id='ddrounding']")
	private static WebElement preferences_RoundDropDown;

	@FindBy(xpath = "//tbody[@id='saveinterest']/tr[1]/td[2]")
	private static WebElement preferences_InterestpermonthRow1WEFField;

	@FindBy(xpath = "//tbody[@id='saveinterest']/tr[1]/td[3]")
	private static WebElement preferences_InterestpermonthRow1InterestField;

	@FindBy(xpath = "//tbody[@id='saveinterest']/tr[1]/td[4]")
	private static WebElement preferences_InterestpermonthRow1DayOfNextMonthField;

	@FindBy(xpath = "//tbody[@id='saveinterest']/tr[2]/td[2]")
	private static WebElement preferences_InterestpermonthRow2WEFField;

	@FindBy(xpath = "//tbody[@id='saveinterest']/tr[2]/td[3]")
	private static WebElement preferences_InterestpermonthRow2InterestField;

	@FindBy(xpath = "//tbody[@id='saveinterest']/tr[2]/td[4]")
	private static WebElement preferences_InterestpermonthRow2DayOfNextMonthField;

	// Sections Elements

	@FindBy(xpath = "//*[@id='divselections']/div/div/div")
	private static WebElement preferences_SectionsBtn;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[1]/td[2])[2]")
	private static WebElement preferences_SectionsRow1TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[1]/td[3])[2]")
	private static WebElement preferences_SectionsRow1TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[1]/td[4])[2]")
	private static WebElement preferences_SectionsRow1SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[2]/td[2])[2]")
	private static WebElement preferences_SectionsRow2TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[2]/td[3])[2]")
	private static WebElement preferences_SectionsRow2TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[2]/td[4])[2]")
	private static WebElement preferences_SectionsRow2SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[3]/td[2])[2]")
	private static WebElement preferences_SectionsRow3TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[3]/td[3])[2]")
	private static WebElement preferences_SectionsRow3TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[3]/td[4])[2]")
	private static WebElement preferences_SectionsRow3SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[4]/td[2])[2]")
	private static WebElement preferences_SectionsRow4TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[4]/td[3])[2]")
	private static WebElement preferences_SectionsRow4TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[4]/td[4])[2]")
	private static WebElement preferences_SectionsRow4SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[5]/td[2])[2]")
	private static WebElement preferences_SectionsRow5TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[5]/td[3])[2]")
	private static WebElement preferences_SectionsRow5TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[5]/td[4])[2]")
	private static WebElement preferences_SectionsRow5SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[6]/td[2])[2]")
	private static WebElement preferences_SectionsRow6TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[6]/td[3])[2]")
	private static WebElement preferences_SectionsRow6TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[6]/td[4])[2]")
	private static WebElement preferences_SectionsRow6SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[7]/td[2])[2]")
	private static WebElement preferences_SectionsRow7TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[7]/td[3])[2]")
	private static WebElement preferences_SectionsRow7TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[7]/td[4])[2]")
	private static WebElement preferences_SectionsRow7SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[8]/td[2])[2]")
	private static WebElement preferences_SectionsRow8TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[8]/td[3])[2]")
	private static WebElement preferences_SectionsRow8TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[8]/td[4])[2]")
	private static WebElement preferences_SectionsRow8SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[9]/td[2])[2]")
	private static WebElement preferences_SectionsRow9TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[9]/td[3])[2]")
	private static WebElement preferences_SectionsRow9TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[9]/td[4])[2]")
	private static WebElement preferences_SectionsRow9SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[10]/td[2])[2]")
	private static WebElement preferences_SectionsRow10TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[10]/td[3])[2]")
	private static WebElement preferences_SectionsRow10TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[10]/td[4])[2]")
	private static WebElement preferences_SectionsRow10SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[11]/td[2])[2]")
	private static WebElement preferences_SectionsRow11TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[11]/td[3])[2]")
	private static WebElement preferences_SectionsRow11TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[11]/td[4])[2]")
	private static WebElement preferences_SectionsRow11SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[12]/td[2])[2]")
	private static WebElement preferences_SectionsRow12TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[12]/td[3])[2]")
	private static WebElement preferences_SectionsRow12TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[12]/td[4])[2]")
	private static WebElement preferences_SectionsRow12SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[13]/td[2])[2]")
	private static WebElement preferences_SectionsRow13TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[13]/td[3])[2]")
	private static WebElement preferences_SectionsRow13TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[13]/td[4])[2]")
	private static WebElement preferences_SectionsRow13SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[14]/td[2])[2]")
	private static WebElement preferences_SectionsRow14TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[14]/td[3])[2]")
	private static WebElement preferences_SectionsRow14TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[14]/td[4])[2]")
	private static WebElement preferences_SectionsRow14SectionCode;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[15]/td[2])[2]")
	private static WebElement preferences_SectionsRow15TdsType;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[15]/td[3])[2]")
	private static WebElement preferences_SectionsRow15TaxDeducted;

	@FindBy(xpath = "(//div[@id='gridsections']/table/tbody/tr[15]/td[4])[2]")
	private static WebElement preferences_SectionsRow15SectionCode;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[16]/td[2]")
	private static WebElement preferences_SectionsRow16TdsType;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[16]/td[3]")
	private static WebElement preferences_SectionsRow16TaxDeducted;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[16]/td[4]")
	private static WebElement preferences_SectionsRow16SectionCode;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[17]/td[2]")
	private static WebElement preferences_SectionsRow17TdsType;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[17]/td[3]")
	private static WebElement preferences_SectionsRow17TaxDeducted;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[17]/td[4]")
	private static WebElement preferences_SectionsRow17SectionCode;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[18]/td[2]")
	private static WebElement preferences_SectionsRow18TdsType;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[18]/td[3]")
	private static WebElement preferences_SectionsRow18TaxDeducted;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[18]/td[4]")
	private static WebElement preferences_SectionsRow18SectionCode;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[19]/td[2]")
	private static WebElement preferences_SectionsRow19TdsType;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[19]/td[3]")
	private static WebElement preferences_SectionsRow19TaxDeducted;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[19]/td[4]")
	private static WebElement preferences_SectionsRow19SectionCode;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[20]/td[2]")
	private static WebElement preferences_SectionsRow20TdsType;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[20]/td[3]")
	private static WebElement preferences_SectionsRow20TaxDeducted;

	@FindBy(xpath = "//div[@id='gridsections']/table/tbody/tr[20]/td[4]")
	private static WebElement preferences_SectionsRow20SectionCode;

	@FindBy(xpath = "//*[@id='ViewSections']/div/div/div[3]/div[2]/div/div[1]/div[1]/span")
	private static WebElement preferences_SectionsAddBtn;

	@FindBy(xpath = "//*[@id='ViewSections']/div/div/div[3]/div[2]/div/div[2]/div[1]/span")
	private static WebElement preferences_SectionsCloseBtn;

	@FindBy(xpath = "//input[@id='txttdstype']")
	private static WebElement preferences_SectionsAddTDStypeTxt;

	@FindBy(xpath = "//input[@id='txtsectioncode']")
	private static WebElement preferences_SectionsAddSectionCodeTxt;

	@FindBy(xpath = "//input[@id='txtsectioncodereturn']")
	private static WebElement preferences_SectionsAddSectionCodeReturnTxt;

	@FindBy(xpath = "//li[@id='btnaddSections']")
	private static WebElement preferences_SectionsAddSectionsSaveBtn;

	@FindBy(xpath = "//li[@id='btnsectionsaddclose']")
	private static WebElement preferences_SectionsAddSectionsCloseBtn;

	// TDS Type Details Elements 1st Row

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[1]/a")
	private static WebElement preferences_TDSTypeDetailsRow1DeleteBtn;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[3]")
	private static WebElement preferences_TDSTypeDetailsRow1NewSectionDropDown;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[4]/input")
	private static WebElement preferences_TDSTypeDetailsRow1TDSTypeAliastxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[5]")
	private static WebElement preferences_TDSTypeDetailsRow1ReasonDropDown;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[6]/input")
	private static WebElement preferences_TDSTypeDetailsRow1NatureOfPaymenttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[7]/input")
	private static WebElement preferences_TDSTypeDetailsRow1NopCodetxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[8]input")
	private static WebElement preferences_TDSTypeDetailsRow1SinglePaymenttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[9]/input")
	private static WebElement preferences_TDSTypeDetailsRow1TotalAmounttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[11]/input")
	private static WebElement preferences_TDSTypeDetailsRow1TDSAccounttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[13]/button")
	private static WebElement preferences_TDSTypeDetailsRow1TaxField;

	/// 2nd Row

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[2]/td[1]/a")
	private static WebElement preferences_TDSTypeDetailsRow2DeleteBtn;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[2]/td[3]")
	private static WebElement preferences_TDSTypeDetailsRow2NewSectionDropDown;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[3]/td[3]/select")
	private static WebElement preferences_TDSTypeDetailsRow3NewSectionDropDown;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[2]/td[4]/input")
	private static WebElement preferences_TDSTypeDetailsRow2TDSTypeAliastxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[2]/td[5]")
	private static WebElement preferences_TDSTypeDetailsRow2ReasonDropDown;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[2]/td[6]/input")
	private static WebElement preferences_TDSTypeDetailsRow2NatureOfPaymenttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[2]/td[7]/input")
	private static WebElement preferences_TDSTypeDetailsRow2NopCodetxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[2]/td[8]input")
	private static WebElement preferences_TDSTypeDetailsRow2SinglePaymenttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[2]/td[9]/input")
	private static WebElement preferences_TDSTypeDetailsRow2TotalAmounttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[3]/td[9]/input")
	private static WebElement preferences_TDSTypeDetailsRow3TotalAmounttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[2]/td[11]/input")
	private static WebElement preferences_TDSTypeDetailsRow2TDSAccounttxt;
	
	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[3]/td[11]/input")
	private static WebElement preferences_TDSTypeDetailsRow3TDSAccounttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[2]/td[13]/button")
	private static WebElement preferences_TDSTypeDetailsRow2TaxField;
	
	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[3]/td[13]/button")
	private static WebElement preferences_TDSTypeDetailsRow3TaxField;

	@FindBy(xpath = "//input[@id='dtpEffectiveDate']")
	private static WebElement preferences_TDSTypeDetailsDefineTaxEffectivedate;

	@FindBy(xpath = "//input[@id='txtPANNotAvailableTDSRate']")
	private static WebElement preferences_TDSTypeDetailsDefineTaxPanNotAvilableTxt;

	@FindBy(xpath = "//input[@id='txtTDSRate']")
	private static WebElement preferences_TDSTypeDetailsDefineTaxTDSRateTxt;

	@FindBy(xpath = "//input[@id='txtSurcharge']")
	private static WebElement preferences_TDSTypeDetailsDefineTaxSurchargeTxt;

	@FindBy(xpath = "//input[@id='txtEducess']")
	private static WebElement preferences_TDSTypeDetailsDefineTaxEduCessTxt;

	@FindBy(xpath = "//*[@id='wndDefineSlab']/div/div/div[3]/div[2]/div/div[1]/div[1]/span")
	private static WebElement preferences_TDSTypeDetailsDefineTaxPreviousTaxsBtn;

	@FindBy(xpath = "//*[@id='btnsaveTaxSlab']/div[1]/span")
	private static WebElement preferences_TDSTypeDetailsDefineTaxSaveBtn;

	@FindBy(xpath = "(//*[@id='btntaxslabclose']/div[1])[1]")
	private static WebElement preferences_TDSTypeDetailsDefineTaxCloseBtn;

	@FindBy(xpath = "//*[@id='PrevtaxslabDetGrid']/tr[1]/td[2]/label")
	private static WebElement preferences_TDSTypeDetailsDefineTaxRow1Datetxt;

	@FindBy(xpath = "//*[@id='PrevtaxslabDetGrid']/tr[1]/td[3]/label")
	private static WebElement preferences_TDSTypeDetailsDefineTaxRow1TDSRatetxt;

	@FindBy(xpath = "//*[@id='PrevtaxslabDetGrid']/tr[1]/td[4]/label")
	private static WebElement preferences_TDSTypeDetailsDefineTaxRow1SurChargetxt;

	@FindBy(xpath = "//*[@id='PrevtaxslabDetGrid']/tr[1]/td[5]/label")
	private static WebElement preferences_TDSTypeDetailsDefineTaxRow1Educesstxt;

	@FindBy(xpath = "//*[@id='wndTDSPreviousTaxesSlab']/div/div/div[3]/div[2]/div/div/div[1]/span")
	private static WebElement preferences_TDSTypeDetailsDefineTaxRow1Close;
	



	// Preferences TCS Elements Start

	@FindBy(xpath = "//button[@id='tab2']")
	private static WebElement preferences_TCSTab;

	@FindBy(xpath = "//input[@id='chkTCSbyTags']")
	private static WebElement preferences_TCSByTagsCheckBox;

	@FindBy(xpath = "//input[@id='rdoTCSNonPublicSector']")
	private static WebElement preferences_TCSNonPublicSectorCheckBox;

	@FindBy(xpath = "//input[@id='rdoTCSPublicSector']")
	private static WebElement preferences_TCSPublicSectorCheckBox;

	@FindBy(xpath = "//input[@id='chkAskforTCSDeduction']")
	private static WebElement preferences_TCSAskForTcsColeectionCheckBox;

	@FindBy(xpath = "//select[@id='TCSByTagsdropdwonId']")
	private static WebElement preferences_TCSByTagsDropDown;

	@FindBy(xpath = "//button[@id='btnTCSVoucherTypeview']")
	private static WebElement preferences_TCSDocumentsTxt;

	@FindBy(xpath = "//select[@id='ddTCSrounding']")
	private static WebElement preferences_TCSRoundingDropDown;

	@FindBy(xpath = "//tbody[@id='TCSinterestbody']/tr[1]/td[2]/input")
	private static WebElement preferences_TCSInterestpermonthRow1WEFField;

	@FindBy(xpath = "//tbody[@id='TCSinterestbody']/tr[1]/td[3]/input")
	private static WebElement preferences_TCSInterestpermonthRow1InteresrText;

	@FindBy(xpath = "//tbody[@id='TCSinterestbody']/tr[1]/td[4]/input")
	private static WebElement preferences_TCSInterestpermonthRow1DayOfNextMonthText;

	// Sections Elements

	@FindBy(xpath = "(//span[@class='glyphicon glyphicon-list'])[2]")
	private static WebElement preferences_TCSSectionsBtn;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[1]/td[2]/label")
	private static WebElement preferences_SectionsRow1TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[1]/td[3]/label")
	private static WebElement preferences_SectionsRow1TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[1]/td[4]/label")
	private static WebElement preferences_SectionsRow1TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[2]/td[2]/label")
	private static WebElement preferences_SectionsRow2TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[2]/td[3]/label")
	private static WebElement preferences_SectionsRow2TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[2]/td[4]/label")
	private static WebElement preferences_SectionsRow2TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[3]/td[2]/label")
	private static WebElement preferences_SectionsRow3TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[3]/td[3]/label")
	private static WebElement preferences_SectionsRow3TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[3]/td[4]/label")
	private static WebElement preferences_SectionsRow3TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[4]/td[2]/label")
	private static WebElement preferences_SectionsRow4TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[4]/td[3]/label")
	private static WebElement preferences_SectionsRow4TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[4]/td[4]/label")
	private static WebElement preferences_SectionsRow4TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[5]/td[2]/label")
	private static WebElement preferences_SectionsRow5TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[5]/td[3]/label")
	private static WebElement preferences_SectionsRow5TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[5]/td[4]/label")
	private static WebElement preferences_SectionsRow5TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[6]/td[2]/label")
	private static WebElement preferences_SectionsRow6TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[6]/td[3]/label")
	private static WebElement preferences_SectionsRow6TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[6]/td[4]/label")
	private static WebElement preferences_SectionsRow6TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[7]/td[2]/label")
	private static WebElement preferences_SectionsRow7TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[7]/td[3]/label")
	private static WebElement preferences_SectionsRow7TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[7]/td[4]/label")
	private static WebElement preferences_SectionsRow7TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[8]/td[2]/label")
	private static WebElement preferences_SectionsRow8TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[8]/td[3]/label")
	private static WebElement preferences_SectionsRow8TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[8]/td[4]/label")
	private static WebElement preferences_SectionsRow8TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[9]/td[2]/label")
	private static WebElement preferences_SectionsRow9TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[9]/td[3]/label")
	private static WebElement preferences_SectionsRow9TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[9]/td[4]/label")
	private static WebElement preferences_SectionsRow9TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[10]/td[2]/label")
	private static WebElement preferences_SectionsRow10TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[10]/td[3]/label")
	private static WebElement preferences_SectionsRow10TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[10]/td[4]/label")
	private static WebElement preferences_SectionsRow10TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[11]/td[2]/label")
	private static WebElement preferences_SectionsRow11TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[11]/td[3]/label")
	private static WebElement preferences_SectionsRow11TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[11]/td[4]/label")
	private static WebElement preferences_SectionsRow11TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[12]/td[2]/label")
	private static WebElement preferences_SectionsRow12TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[12]/td[3]/label")
	private static WebElement preferences_SectionsRow12TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[12]/td[4]/label")
	private static WebElement preferences_SectionsRow12TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[13]/td[2]/label")
	private static WebElement preferences_SectionsRow13TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[13]/td[3]/label")
	private static WebElement preferences_SectionsRow13TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[13]/td[4]/label")
	private static WebElement preferences_SectionsRow13TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[14]/td[2]/label")
	private static WebElement preferences_SectionsRow14TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[14]/td[3]/label")
	private static WebElement preferences_SectionsRow14TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[14]/td[4]/label")
	private static WebElement preferences_SectionsRow14TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[15]/td[2]/label")
	private static WebElement preferences_SectionsRow15TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[15]/td[3]/label")
	private static WebElement preferences_SectionsRow15TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[15]/td[4]/label")
	private static WebElement preferences_SectionsRow15TcsSectionCode;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[16]/td[2]/label")
	private static WebElement preferences_SectionsRow16TcsType;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[16]/td[3]/label")
	private static WebElement preferences_SectionsRow16TcsTaxDeducted;

	@FindBy(xpath = "//tbody[@id='TCSSectionsBody']/tr[16]/td[4]/label")
	private static WebElement preferences_SectionsRow16TcsSectionCode;

	@FindBy(xpath = "//*[@id='ViewTCSSections']/div/div/div[3]/div[2]/div/div[1]")
	private static WebElement preferences_SectionsTcsAddBtn;

	@FindBy(xpath = "//*[@id='ViewTCSSections']/div/div/div[3]/div[2]/div/div[2]")
	private static WebElement preferences_SectionsTcsCloseBtn;

	// Sections Add Button Elements

	@FindBy(xpath = "//input[@id='txtTCStype']")
	private static WebElement preferences_SectionsAddTCStypeTxt;

	@FindBy(xpath = "//input[@id='txtTCSsectioncode']")
	private static WebElement preferences_SectionsAddTCSSectionCodeTxt;

	@FindBy(xpath = "//input[@id='txtTCSsectioncodereturn']")
	private static WebElement preferences_SectionsAddTCSSectionCodeReturnTxt;

	@FindBy(xpath = "//div[@id='btnaddTCSSections']")
	private static WebElement preferences_SectionsAddTCSSectionsSaveBtn;

	@FindBy(xpath = "//*[@id='AddTCSSections']/div/div/div[3]/div[2]/div/div[2]")
	private static WebElement preferences_SectionsAddTCSSectionsCloseBtn;

	// TCS TYPE DETAILS ELEMENTS ROW1

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[1]/td[3]/select")
	private static WebElement preferences_TCSTypeDetailsRow1NewSectionDropDown;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[1]/td[4]/input")
	private static WebElement preferences_TDSTypeDetailsRow1TCSTypeAliastxt;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[1]/td[5]/select")
	private static WebElement preferences_TDSTypeDetailsRow1TCSReasonDropDown;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[6]/input")
	private static WebElement preferences_TCSTypeDetailsRow1NatureOfPaymenttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[7]/input")
	private static WebElement preferences_TDSTypeDetailsRow1TCSNopCodetxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[8]input")
	private static WebElement preferences_TDSTypeDetailsRow1TCSSinglePaymenttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[9]/input")
	private static WebElement preferences_TDSTypeDetailsRow1TCSTotalAmounttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[11]/input")
	private static WebElement preferences_TDSTypeDetailsRow1TCSAccounttxt;

	@FindBy(xpath = "//tbody[@id='saveTDSDet']/tr[1]/td[13]/button")
	private static WebElement preferences_TDSTypeDetailsRow1TCSTaxField;

	// TCS TYPE DETAILS ELEMENTS ROW2

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[2]/td[3]/select")
	private static WebElement preferences_TCSTypeDetailsRow2NewSectionDropDown;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[2]/td[4]/input")
	private static WebElement preferences_TCSTypeDetailsRow2TCSTypeAliastxt;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[2]/td[5]/select")
	private static WebElement preferences_TCSTypeDetailsRow2TCSReasonDropDown;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[2]/td[6]/input")
	private static WebElement preferences_TCSTypeDetailsRow2NatureOfPaymenttxt;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[2]/td[7]/input")
	private static WebElement preferences_TCSTypeDetailsRow2TCSNopCodetxt;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[2]/td[8]input")
	private static WebElement preferences_TCSTypeDetailsRow2TCSSinglePaymenttxt;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[2]/td[9]/input")
	private static WebElement preferences_TCSTypeDetailsRow2TCSTotalAmounttxt;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[2]/td[11]/input")
	private static WebElement preferences_TCSTypeDetailsRow2TCSAccounttxt;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[3]/td[11]/input")
	private static WebElement preferences_TCSTypeDetailsRow3TCSAccounttxt;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[2]/td[13]/button")
	private static WebElement preferences_TCSTypeDetailsRow2TCSTaxField;

	@FindBy(xpath = "//tbody[@id='saveTCSDet']/tr[3]/td[13]/button")
	private static WebElement preferences_TCSTypeDetailsRow3TCSTaxField;

	@FindBy(xpath = "//input[@id='dtpTCSEffectiveDate']")
	private static WebElement preferences_TCSTypeDetailsDefineTaxEffectivedate;

	@FindBy(xpath = "//input[@id='txtPANNotAvailableTCSRate']")   
	private static WebElement preferences_TCSTypeDetailsDefineTaxPanNotAvilableTxt;

	@FindBy(xpath = "//input[@id='txtTCSRate']")
	private static WebElement preferences_TCSTypeDetailsDefineTaxTCSRateTxt;

	@FindBy(xpath = "//input[@id='txtTCSSurcharge']")
	private static WebElement preferences_TCSTypeDetailsDefineTaxSurchargeTxt;

	@FindBy(xpath = "//input[@id='txtTCSEducess']")
	private static WebElement preferences_TCSTypeDetailsDefineTaxEduCessTxt;

	@FindBy(xpath = "//*[@id='DefineSlabPop']/div/div/div[3]/div[2]/div/div[1]")
	private static WebElement preferences_TCSTypeDetailsDefineTaxPreviousTaxsBtn;

	@FindBy(xpath = "//div[@id='btnsaveTCSTaxSlab']")
	private static WebElement preferences_TCSTypeDetailsDefineTaxSaveBtn;

	@FindBy(xpath = "//*[@id='DefineSlabPop']/div/div/div[3]/div[2]/div/div[3]")
	private static WebElement preferences_TCSTypeDetailsDefineTaxCloseBtn;

	// Previous Tax's Elements

	@FindBy(xpath = "//*[@id='TCSPrevtaxslabDetGrid']/tr/td[2]/label")
	private static WebElement preferences_TCSTypeDetailsDefineTaxRow1Datetxt;

	@FindBy(xpath = "//*[@id='TCSPrevtaxslabDetGrid']/tr/td[3]/label")
	private static WebElement preferences_TCSTypeDetailsDefineTaxRow1TDSRatetxt;

	@FindBy(xpath = "//*[@id='TCSPrevtaxslabDetGrid']/tr/td[4]/label")
	private static WebElement preferences_TCSTypeDetailsDefineTaxRow1SurChargetxt;

	@FindBy(xpath = "//*[@id='TCSPrevtaxslabDetGrid']/tr/td[5]/label")
	private static WebElement preferences_TCSTypeDetailsDefineTaxRow1Educesstxt;

	@FindBy(xpath = "//*[@id='TCSPreviousTaxesSlab']/div/div/div[3]/div[2]/div/div")
	private static WebElement preferences_TCSTypeDetailsDefineTaxRow1Close;

	@FindBy(xpath = "//a[@id='60']/div")
	private static WebElement financialsMenu;
	
	@FindBy(xpath = "//a[@id='2001']/span")
	private static WebElement cashAndBankMenu;
	
	@FindBy(xpath = "//a[@id='2002']/span")
	private static WebElement receiptsMenu;
	
	@FindBy(xpath = "//input[@id='id_header_4']")
	private static WebElement cashAnsBankActTxt;


	@FindBy(xpath = "//a[@id='4501']/span")
	private static WebElement statutoryMenu;
	
	@FindBy(xpath = "//a[@id='2023']/span")
	private static WebElement journalsMenu;
	
	@FindBy(xpath = "//a[@id='2052']/span")
	private static WebElement tcsJVMenu;
	
	@FindBy(xpath = "//input[@id='id_header_1']")
	private static WebElement documentNoTxt;
	
	@FindBy(xpath = "//input[@id='id_header_67108967']")
	private static WebElement tcsVocherno;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td")
	 private static List<WebElement> jsVocherBodyRow1List;
	 
	 @FindBy(xpath="//div[@id='id_transactionentry_close']")
	 private static WebElement tcsJvCloseBtn;
	 
	
	 @FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr/td[2]/input")
		private static List<WebElement> grid_CheckBoxList;
		
	@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr/td[5]")
		private static List<WebElement> grid_VoucherNoList;
	
	@FindBy(xpath = "//a[@id='2056']/span")
	private static WebElement salesDebitNotesMenu;
	
	@FindBy(xpath = "//div[@id='id_transaction_homescreen_new']")
	private static WebElement new_newBtn;
	
	@FindBy(xpath = "//input[@id='id_header_4']")
	private static WebElement vendorAccounttxt;
	
	@FindBy(xpath = "//input[@id='id_body_12']")
	private static WebElement enter_AccountTxt;
	
	@FindBy(xpath = "//input[@id='id_body_16']")
	private static WebElement enter_AmountTxt;
	
	@FindBy(xpath = "//td[contains(text(), 'SalInv:1')]//../td[2]/input")
	private static WebElement referenceGridCheckBox;
	
	
	@FindBy(xpath = "//li[@id='btn_Export']")
	private static WebElement transactionDetailsExportBtn;
	
	// TCS Challan Elements

		

		@FindBy(xpath = "//div[@id='btndelete']")
		private static WebElement tcsChallanDeleteBtn;

		@FindBy(xpath = "//*[@id='TDSChallandiv']/div[1]/div[2]/div/div[3]")
		private static WebElement tcsChallanCloseBtn;

		@FindBy(xpath = "//select[@id='popTDSserialnoId']")
		private static WebElement tcsChallanSerialNoDropDown;

		
		@FindBy(xpath = "//select[@id='ddlchallantypeId']")
		private static WebElement tcsChallanChallanTypeDropDown;

		@FindBy(xpath = "//input[@id='txtchlnNo']")
		private static WebElement tcsChallanChallanNoTxt;

		@FindBy(xpath = "//input[@id='txtdateofchallan']")
		private static WebElement tcsChallanDateOfChallanTxt;

		@FindBy(xpath = "//select[@id='ddlchallanmonthId']")
		private static WebElement tcsChallanMonthDropDown;

		@FindBy(xpath = "//select[@id='ddlbanknameId']")
		private static WebElement tcsChallanBankNameDropDown;

		@FindBy(xpath = "//select[@id='ddlPaymentModeId']")
		private static WebElement tcsChallanPaymentModeDropDown;

		@FindBy(xpath = "//input[@id='txtChkNo']")
		private static WebElement tcsChallanChequeNoTxt;

		@FindBy(xpath = "//input[@id='txtdate']")
		private static WebElement tcsChallanDateTxt;

		@FindBy(xpath = "//select[@id='ddlDrawnonId']")
		private static WebElement tcsChallanDrawnOnDropDown;

		@FindBy(xpath = "//input[@id='txtincometax']")
		private static WebElement tcsChallanIncomeTaxTxt;

		@FindBy(xpath = "//input[@id='txtsurcharge']")
		private static WebElement tcsChallanSurchargeTxt;

		@FindBy(xpath = "//input[@id='txteccess']")
		private static WebElement tcsChallanEducationCessTxt;

		@FindBy(xpath = "//input[@id='txtinterest']")
		private static WebElement tcsChallanInterestTxt;

		@FindBy(xpath = "//input[@id='txtpenalty']")
		private static WebElement tcsChallanpenaltyTxt;

		@FindBy(xpath = "//input[@id='txtothers']")
		private static WebElement tcsChallanOthersTxt;

		@FindBy(xpath = "//input[@id='txttotalamount']")
		private static WebElement tcsChallanTotalAmountTxt;

		@FindBy(xpath = "//*[@id='btnCalc']/i")
		private static WebElement tcsChallanInterestSideBtn;
	
	@FindBy(xpath = "//div[@id='jconfirm-box78967']/div")
	private static WebElement ExcelConfirmTxt;
	
	@FindBy(xpath = "//i[@class='icon-accounts icon-font6']")
	private static WebElement transactionDetailsStAllYesOrNo;
	
	@FindBy(xpath = "//li[@id='btnGenChalan']")
	private static WebElement transactionDetailsGenerateChallan;
	
	@FindBy(xpath = "//span[@class='icon-save data-accept']")
	private static WebElement tcsChallnCloseBtn;
	
	@FindBy(xpath = "(//button[@class='btn btn-default'])[1]")
	private static WebElement ExcelConfirmYesBtn;
	
	@FindBy(xpath="//*[@id='divErrorAlert']/div/table/tbody/tr/td[2]/div[2]")
	private static WebElement errorMessageTxt;

	@FindBy(xpath="//*[@id='divErrorAlert']/div/table/tbody/tr/td[3]/span")
	private static WebElement errorMessageClose;
	
	@FindBy(xpath = "//*[@id='tdsmultiselect']/div/button")
	private static WebElement tcsChallanTcsTypetxt;

	
	
	
	@FindBy(xpath="//input[@id='txtInvoiceNo0']")
	private static WebElement postTax_InvoiceNoTxt;

@FindBy(xpath="//div[@id='CrDrGSTTaxPop']/div/div[2]/div[2]/button[1]")
	private static WebElement postTax_PostBtn;

	
	@FindBy(xpath = "//iframe[@id='id_global_externalmodule_container']")
	private static WebElement MainFrame;

	@FindBy(xpath = "//a[@id='4502']/span")
	private static WebElement tdsTcsMenu;  
	
	@FindBy(xpath = "//a[@id='4517']/span")
	private static WebElement tcsChallanMenu;

	@FindBy(xpath = "//a[@id='4504']/span")
	private static WebElement preferencesMenu;

	@FindBy(xpath = "//section[@id='page_Content']/iframe")
	private static WebElement iframeElement;

	@FindBy(xpath = "//div[@class='theme_color font-6']")
	public static WebElement errorMessage;

	@FindBy(xpath = "//span[@class='icon-reject2 theme_color']") // span[@class='icon-reject2 theme_color']
	public static WebElement errorMessageCloseBtn;

	// Define Tax Slab Elements

	@FindBy(xpath = "//*[@id='taxslabDetGrid']/tr[1]/td[4]/input")
	public static WebElement defineTaxRow1SlabFrom;

	@FindBy(xpath = "//*[@id='taxslabDetGrid']/tr[1]/td[5]/input")
	public static WebElement defineTaxRow1SlabTo;

	@FindBy(xpath = "//*[@id='taxslabDetGrid']/tr[1]/td[6]/input")
	public static WebElement defineTaxRow1SlabTdsRate;

	@FindBy(xpath = "//*[@id='taxslabDetGrid']/tr[1]/td[8]/input")
	public static WebElement defineTaxRow1SlabEducationCess;

	@FindBy(xpath = "//*[@id='taxslabDetGrid']/tr[2]/td[4]/input")
	public static WebElement defineTaxRow2SlabFrom;

	@FindBy(xpath = "//*[@id='taxslabDetGrid']/tr[2]/td[5]/input")
	public static WebElement defineTaxRow2SlabTo;

	@FindBy(xpath = "//*[@id='taxslabDetGrid']/tr[2]/td[6]/input")
	public static WebElement defineTaxRow2SlabTdsRate;

	@FindBy(xpath = "//div[@id='btnsavedefineTaxSlab']")
	public static WebElement defineTaxSlabSaveBtn;

	@FindBy(xpath = "(//div[@id='btntaxslabclose'])[2]")
	public static WebElement defineTaxSlabCloseBtn;

	@FindBy(xpath = "//*[@id='wndDefineTaxSlab']/div/div/div[3]/div[2]/div/div[1]")
	public static WebElement defineTaxSlabPriviousTaxBtn;

	@FindBy(xpath = "//input[@id='txtslabPANNotAvailableTDSRate']")
	public static WebElement defineTaxSlabPANNotAvailableTdsRate;
	
	//Sales vocher elements
	
	@FindBy(xpath="//span[@id='btnPick']")
	private static WebElement  billRefPickIcon;		

	@FindBy(xpath="//span[@id='btnOk']")
	private static WebElement  billRefOkBtn;		

	@FindBy(xpath="//span[@id='btnCancel']")
	private static WebElement  billRefcancel;	

	@FindBy(xpath="//div[@id='myNavbar']//span[@class='icon-new icon-font6']")
	private static WebElement  newBtn;



@FindBy(xpath="//*[@id='61']/span")
    private static WebElement  financialsTransactionMenu; 
@FindBy(xpath="//a[@id='2014']/span")
    private static WebElement financialsTransactionsSalesMenu;

@FindBy(xpath="//a[@id='2015']/span")
    private static WebElement SalesInvoiceVoucher;

@FindBy(xpath="//*[@id='id_header_3']")
    private static WebElement salesAccountTxt;

@FindBy(xpath="//input[@id='id_header_67110203']")
private static WebElement tcsBillAmmountTxt;

@FindBy(xpath="//input[@id='id_footer_134217771']")
private static WebElement tcsTxt;


@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_table_body']/tr[1]/td")
private static List<WebElement> voucherBodyRow1List; 

    @FindBy(xpath="//*[@id='id_header_4']")
    private static WebElement CustomerAccounttxt;

@FindBy(xpath="//input[@id='id_header_67108922']")
    private static WebElement  salInvPlaceOfSupplyTxt;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[2]")
    private static WebElement  select1stRow_1stColumn;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[3]")
    private static WebElement  select1stRow_2ndColumn;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[4]")
    private static WebElement  select1stRow_3rdColumn;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[5]")
    private static WebElement  select1stRow_4thColumn;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[6]")
    private static WebElement  select1stRow_5thColumn;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[7]")
    private static WebElement  select1stRow_6thColumn;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[8]")
    private static WebElement  select1stRow_7thColumn;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[9]")
    private static WebElement  select1stRow_8thColumn;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[10]")
    private static WebElement  select1stRow_9thColumn;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[11]")
    private static WebElement  select1stRow_10thColumn;

    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[12]")
    private static WebElement  select1stRow_11thColumn;

    @FindBy(xpath="//*[@id='id_body_23']")
    private static WebElement  enter_ItemTxt;

    @FindBy(xpath="//*[@id='id_body_26']")
    private static WebElement  enter_QuantityTxt;

    @FindBy(xpath="//*[@id='id_body_27']")
    private static WebElement  enter_RateTxt;

    @FindBy(xpath="//*[@id='id_body_28']")
    private static WebElement  enter_GrossTxt;

    @FindBy(xpath="//*[@id='id_body_33555529']")
    private static WebElement  enter_CGSTTxt;

    @FindBy(xpath="//*[@id='id_body_33554479']")
    private static WebElement enter_PR_CGSTTxt;

    @FindBy(xpath="//*[@id='id_body_33555530']")
    private static WebElement  enter_SGSTTxt;

    @FindBy(xpath="//*[@id='id_body_33554480']")
    private static WebElement  enter_PR_SGSTTxt;

    @FindBy(xpath="//*[@id='id_body_33555531']")
    private static WebElement  enter_IGSTTxt;

    @FindBy(xpath="//*[@id='id_body_33554481']")
    private static WebElement  enter_PR_IGSTTxt;

    @FindBy(xpath="//*[@id='id_body_33555532']")
    private static WebElement  enter_CessTxt;

    @FindBy(xpath="//*[@id='id_body_33554482']")
    private static WebElement  enter_PR_CessTxt;

    @FindBy(xpath="//*[@id='id_body_33555533']")
    private static WebElement  enter_TaxableValue;

    @FindBy(xpath="//*[@id='id_body_33554483']")
    private static WebElement  enter_PR_TaxableValue;

    @FindBy(xpath="//select[@id='id_body_16778395']")
    private static WebElement  enter_TypeOfPurchaseDropdown;

    @FindBy(xpath="//*[@id='id_transactionentry_save']")
    private static WebElement VoucherSaveBtn;

    @FindBy(xpath="//*[@id='id_body_33554455']")
    private static WebElement  enter_SalInvCGSTTxt;
    
    @FindBy(xpath="//*[@id='id_body_33554456']")
    private static WebElement  enter_SalInvSGSTTxt;
    
    @FindBy(xpath="//*[@id='id_body_33554457']")
    private static WebElement  enter_SalInvIGSTTxt;
    
    @FindBy(xpath="//*[@id='id_body_33554458']")
    private static WebElement  enter_SalInvCessTxt;

    @FindBy(xpath="//*[@id='id_body_33554459']")
    private static WebElement  enter_SalInvTaxableValueTxt;
    
    @FindBy(xpath="//iframe[@id='id_global_externalmodule_container']")
    private static WebElement GSTVoucherValidationFrame;
    
    @FindBy(xpath="//iframe[@id='id_global_externalmodule_container']")
	private static WebElement  TCSFrame;
	
	@FindBy(xpath="(//select[@id='ddTDSType']/option)[1]")
	private static WebElement  TCSTypeDropdownSelectedOpn;
	
	@FindBy(xpath="//select[@id='ddTDSType']")
	private static WebElement  TCSTypeDropdown;
	
	
	
	
	@FindBy(xpath="//*[@id='divbtns']/button")
	private static WebElement  TCSPostBtn;
	
	 @FindBy(xpath="//input[@id='txtTDSSection']")
	 private static WebElement TCSDeductionPopupTDSSectionTxt;
	 
	 @FindBy(xpath="//input[@id='txtAmountofpayment']")
	 private static WebElement TCSDeductionPopupAmountTxt;
	 
	 @FindBy(xpath="//input[@id='txtIncTaxper']")
	 private static WebElement TCSDeductionPopupTaxPercentageTxt;
	 
	 @FindBy(xpath="//input[@id='txtIncTax']")
	 private static WebElement TCSDeductionPopupTaxAmountTxt;
	 
	 @FindBy(xpath="//input[@id='txtTaxDeducted']")
	 private static WebElement TCSDeductionPopupTaxDeductedTxt;
	 
	
	 @FindBy(xpath="//input[@id='btnpopclose']")
	 private static WebElement TCSDeductionPopupCancelBtn;
	 
	 @FindBy(xpath="//*[@id='id_transactionentry_previous']/div/span")
		private static WebElement VoucherPreviousBtn;
	 
	 
	 // Vendor Tcs Elements
	 
		@FindBy(xpath = "//*[@id='1']/div/span")
		public static WebElement homeMenu;

		// Masters Menu
		@FindBy(xpath = "//*[@id='1000']/span")
		public static WebElement mastersMenu;

		// Accounts
		@FindBy(xpath = "//a[@id='1104']//span[contains(text(),'Account')]")
		public static WebElement accounts;

		@FindBy(xpath = "//*[@id='cmbUserTypeMaster']")
		public static WebElement accountsSearchTxt;

		@FindBy(xpath = "//*[@id='btnSearchAcc1']/i")
		public static WebElement accountsSearchIcon;

		@FindBy(xpath = "//*[@id='LandingGridBody']/tr/td[8]//input")
		private static List<WebElement> masterGridBodyChkbox;

		@FindBy(xpath = "//*[@id='LandingGridBody']/tr/td[11]")
		private static List<WebElement> masterGridBodyName;

		@FindBy(xpath = "//*[@id='divBtnGroup6']/div/a")
		private static WebElement Accounts_TCSTab;

		@FindBy(xpath = "//tbody[@id='grdCVTCSTypeDetails']//tr//td[7]")
		private static List<WebElement> TCSTypeNameList;

		

		@FindBy(xpath = "//tbody[@id='grdCVTCSTypeDetails']//tr//td//input[@class='TCScheckBoxGrid']")
		private static List<WebElement> TCSTypeChecklist1;

		@FindBy(xpath = "//*[@id='btnClose']/i")
		private static WebElement Accounts_CloseBtn;

		@FindBy(xpath = "//*[@id='btnMasterSaveClick']")
		public static WebElement saveBtn;

		@FindBy(xpath = "//iframe[@id='imyFrame6']")
		public static WebElement tcsfRame;	
	
    

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

	public static boolean checkAvailablityOfTDSTab() throws InterruptedException {

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(statutoryMenu));
		statutoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tdsTcsMenu));
		tdsTcsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesMenu));
		preferencesMenu.click();

		getFrame(iframeElement);

		Thread.sleep(3000);

		boolean actLTdsTab = preferences_TDSTab.isDisplayed();

		boolean expTdsTab = true;

		if (!(preferences_AskForTdsDeductionCheckBox.isSelected())) {
			preferences_AskForTdsDeductionCheckBox.click();
			// System.out.println("enabled");
		}

		if (!(preferences_DefineSlabCheckBox.isSelected())) {
			preferences_DefineSlabCheckBox.click();
			// System.out.println("enabled");

		}
		if (preferences_AllowMultipleTdsCheckBox.isSelected()) {
			preferences_AllowMultipleTdsCheckBox.click();
			// System.out.println("disabled");
		}

		System.out.println(
				"********************************** checkAvailablityOfTDSTab  *******************************************************");

		System.out.println("Check actLTdsTab     	:" + actLTdsTab);
		System.out.println("Check expTdsTab     	:" + expTdsTab);

		if (actLTdsTab == expTdsTab) {
			return true;
		} else {
			return false;
		}

	}

	public static boolean clickOnVocherTypeAndvalidateTheVochers()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferences_VocherType));
		preferences_VocherType.click();

		boolean actLRow1CheckBox = preferences_VocherTypeRow1CheckBox.isSelected();
		boolean actLRow3CheckBox = preferences_VocherTypeRow3CheckBox.isSelected();
		boolean actLRow8CheckBox = preferences_VocherTypeRow8CheckBox.isSelected();
		boolean actLRow9CheckBox = preferences_VocherTypeRow9CheckBox.isSelected();

		boolean expRow1CheckBox = true;
		boolean expRow3CheckBox = true;
		boolean expRow8CheckBox = true;
		boolean expRow9CheckBox = true;

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferences_TDSVocherTypeOkBtn));
		preferences_TDSVocherTypeOkBtn.click();

		checkValidationMessage("Data Saved Successfully");

		System.out.println(
				"********************************** clickOnVocherTypeAndvalidateTheVochers  *******************************************************");

		System.out.println("actLRow1CheckBox     	:" + actLRow1CheckBox);
		System.out.println("expRow1CheckBox	    	:" + expRow1CheckBox);

		System.out.println("actLRow3CheckBox     	:" + actLRow3CheckBox);
		System.out.println("expRow3CheckBox     	:" + expRow3CheckBox);

		System.out.println("actLRow8CheckBox     	:" + actLRow8CheckBox);
		System.out.println("expRow8CheckBox     	:" + expRow8CheckBox);

		System.out.println("actLRow9CheckBox     	:" + actLRow9CheckBox);
		System.out.println("expRow9CheckBox     	:" + expRow9CheckBox);

		if (actLRow1CheckBox == expRow1CheckBox && actLRow3CheckBox == expRow3CheckBox
				&& actLRow8CheckBox == expRow8CheckBox && actLRow9CheckBox == expRow9CheckBox) {
			return true;
		} else {
			return false;
		}
	}

	public static boolean enterTdsTypeDetails()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(preferences_TDSTypeDetailsRow3NewSectionDropDown));
		preferences_TDSTypeDetailsRow3NewSectionDropDown.click();
		preferences_TDSTypeDetailsRow3NewSectionDropDown.sendKeys("Insurance Commission");

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(preferences_TDSTypeDetailsRow3TotalAmounttxt));
		preferences_TDSTypeDetailsRow3TotalAmounttxt.click();

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(preferences_TDSTypeDetailsRow3TDSAccounttxt));
		preferences_TDSTypeDetailsRow3TDSAccounttxt.click();
		preferences_TDSTypeDetailsRow3TDSAccounttxt.sendKeys("Authorized Capital");

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(preferences_TDSTypeDetailsRow3TaxField));
		preferences_TDSTypeDetailsRow3TaxField.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineTaxRow1SlabFrom));
		defineTaxRow1SlabFrom.sendKeys("1");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineTaxRow1SlabTo));
		defineTaxRow1SlabTo.sendKeys("20000");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineTaxRow1SlabTdsRate));
		defineTaxRow1SlabTdsRate.sendKeys("5");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineTaxRow1SlabEducationCess));
		defineTaxRow1SlabEducationCess.click();
		defineTaxRow1SlabEducationCess.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineTaxRow2SlabFrom));
		defineTaxRow2SlabFrom.sendKeys("20001");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineTaxRow2SlabTo));
		defineTaxRow2SlabTo.sendKeys("35000");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineTaxRow2SlabTdsRate));
		defineTaxRow2SlabTdsRate.sendKeys("10");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineTaxSlabSaveBtn));
		defineTaxSlabSaveBtn.click();

		checkValidationMessage("PAN Not Available TDS Rate is mandatory");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineTaxSlabPANNotAvailableTdsRate));
		defineTaxSlabPANNotAvailableTdsRate.sendKeys("10");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineTaxSlabSaveBtn));
		defineTaxSlabSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferences_SaveBtn));
		preferences_SaveBtn.click();

		String expMessage1 = "Data Saved Successfully";
		String actMessage = checkValidationMessage(expMessage1);
		System.out.println("Actual Message     :  " + actMessage);
		System.out.println("exp Message        :  " + expMessage1);

		if (actMessage.equalsIgnoreCase(expMessage1)) {
			return true;
		} else {
			return false;
		}

	}
	
	public static boolean checkAvailablityOfTCSTab() throws InterruptedException {

		
		click(financialsMenu);
		
		click(statutoryMenu);	
		
		click(tdsTcsMenu);	
		
		click(preferencesMenu);

		getFrame(iframeElement);

		Thread.sleep(3000);

		boolean actLTcsTab = preferences_TCSTab.isDisplayed();

		boolean expTcsTab = true;		
		
		click(preferences_TCSTab);
		
		Thread.sleep(3000);		

		System.out.println(
				"********************************** checkAvailablityOfTCSTab  *******************************************************");

		System.out.println("Check actLTdsTab     	:" + actLTcsTab);
		System.out.println("Check expTdsTab     	:" + expTcsTab);
		
		if (actLTcsTab == (expTcsTab)) {
			return true;
		} else {
			return false;
		}
	}
	
	public static boolean clickOnVocherTypeAndvalidateTheVochersInTCS()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {

		
		click(preferences_TCSVocherType);
		
		Thread.sleep(2000);

		boolean actLRow1CheckBox = preferences_TcsVocherTypeRow1CheckBox.isSelected();
		boolean actLRow4CheckBox = preferences_TcsVocherTypeRow4CheckBox.isSelected();
		boolean actLRow5CheckBox = preferences_TcsVocherTypeRow5CheckBox.isSelected();
		

		boolean expRow1CheckBox = true;
		boolean expRow4CheckBox = true;
		boolean expRow5CheckBox = true;
		
		click(preferences_TcsVocherTypeRow1FieldsBtn);
		
		Thread.sleep(2000);
		
		String row1actLFieldName    =Row1editScreenAndFooterFieldsName.getText();
		String row1expFieldname		="Discount";
				
		String row2actLFieldName    =Row2editScreenAndFooterFieldsName.getText();
		String row2expFieldname		="CGST";
		
		String row3actLFieldName    =Row3editScreenAndFooterFieldsName.getText();
		String row3expFieldname		="SGST";
		
		String row4actLFieldName    =Row4editScreenAndFooterFieldsName.getText();
		String row4expFieldname		="IGST";
		
		String row5actLFieldName    =Row5editScreenAndFooterFieldsName.getText();
		String row5expFieldname		="Cess";
		
		String row6actLFieldName    =Row6editScreenAndFooterFieldsName.getText();
		String row6expFieldname		="Taxable Value";
		
		click(editScreenAndFooterOkBtn);
		
		Thread.sleep(2000);
		
		click(tcsVocherTypeOkBtn);
		
		
		checkValidationMessage("Data Saved Successfully");
		

		System.out.println(
				"********************************** clickOnTCSVocherTypeAndvalidateTheVochers  *******************************************************");

		System.out.println("actLRow1CheckBox     	:" + actLRow1CheckBox);
		System.out.println("expRow1CheckBox	    	:" + expRow1CheckBox);

		System.out.println("actLRow3CheckBox     	:" + actLRow4CheckBox);
		System.out.println("expRow3CheckBox     	:" + expRow4CheckBox);

		System.out.println("actLRow8CheckBox     	:" + actLRow5CheckBox);
		System.out.println("expRow8CheckBox     	:" + expRow5CheckBox);
		
		System.out.println("row1actLFieldName     	:" + row1actLFieldName);
		System.out.println("row1actLFieldName     	:" + row1actLFieldName);
		
		System.out.println("row2actLFieldName     	:" + row2actLFieldName);
		System.out.println("row2expFieldname	    :" + row2expFieldname);
		
		System.out.println("row3actLFieldName     	:" + row3actLFieldName);
		System.out.println("row3expFieldname	    :" + row3expFieldname);
		
		System.out.println("row4actLFieldName     	:" + row4actLFieldName);
		System.out.println("row4expFieldname	    :" + row4expFieldname);
		
		System.out.println("row5actLFieldName     	:" + row5actLFieldName);
		System.out.println("row5expFieldname	    :" + row5expFieldname);
		
		System.out.println("row6actLFieldName     	:" + row6actLFieldName);
		System.out.println("row6expFieldname	    :" + row6expFieldname);

		
		

		if (actLRow1CheckBox == expRow1CheckBox && actLRow4CheckBox == expRow4CheckBox && actLRow5CheckBox == expRow5CheckBox 
				&& row1actLFieldName.equals(row1expFieldname) && row2actLFieldName.equalsIgnoreCase(row2expFieldname) && 
						row3actLFieldName.equalsIgnoreCase(row3expFieldname) &&  row4actLFieldName.equalsIgnoreCase(row4expFieldname) && 
								row5actLFieldName.equalsIgnoreCase(row5expFieldname) &&  row6actLFieldName.equalsIgnoreCase(row6expFieldname)) {
			return true;
		} else {
			return false;
		}
	}
	
	public static boolean enterTcsTypeDetails()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		
		Thread.sleep(2000);
		
		click(preferences_TCSTypeDetailsRow2NewSectionDropDown);
		enterText(preferences_TCSTypeDetailsRow2NewSectionDropDown, "Sale in cash of any goods (other than bullion/jewelry)");
		
		click(preferences_TCSTypeDetailsRow2TCSTotalAmounttxt);
		
		click(preferences_TCSTypeDetailsRow2TCSAccounttxt);
		enterText(preferences_TCSTypeDetailsRow2TCSAccounttxt, "Authorized Capital");
		
		click(preferences_TCSTypeDetailsRow2TCSTaxField);

		Thread.sleep(3000);

		
		enterText(preferences_TCSTypeDetailsDefineTaxTCSRateTxt, "7");

		
		enterText(preferences_TCSTypeDetailsDefineTaxSurchargeTxt, "0");

		
		enterText(preferences_TCSTypeDetailsDefineTaxEduCessTxt, "0");

		
		click(preferences_TCSTypeDetailsDefineTaxSaveBtn);
		

		checkValidationMessage("PAN Not Available TDS Rate is mandatory");
		
		
		enterText(preferences_TCSTypeDetailsDefineTaxPanNotAvilableTxt, "5");

		
		click(preferences_TCSTypeDetailsDefineTaxSaveBtn);
		
		click(preferences_SaveBtn);

		String expMessage1 = "Data Saved Successfully";
		String actMessage = checkValidationMessage(expMessage1);
		System.out.println("Actual Message     :  " + actMessage);
		System.out.println("exp Message        :  " + expMessage1);

		if (actMessage.equalsIgnoreCase(expMessage1)) {
			return true;
		} else {
			return false;
		}

	}
	
	 public static boolean checkValidateEnableTCSTypeInCustomerTCSAccount() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
			getDriver().navigate().refresh();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
			mastersMenu.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
			accounts.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountsSearchTxt));
			accountsSearchTxt.click();
			accountsSearchTxt.sendKeys("Customer TCS");
			Thread.sleep(3000);
			accountsSearchTxt.sendKeys(Keys.TAB);

			int count = masterGridBodyName.size();

			ArrayList<String> AccountNames = new ArrayList<String>();

			for (int i = 0; i < count; i++) {
				String data = masterGridBodyName.get(i).getText();
				AccountNames.add(data);

				if (data.equalsIgnoreCase("Customer TCS")) {
					getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
					break;
				}
			}

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Accounts_TCSTab));
			Accounts_TCSTab.click();

			Thread.sleep(5000);

			getDriver().switchTo().frame(tcsfRame);
			Thread.sleep(2000);

			int count1 = TCSTypeNameList.size();

			ArrayList<String> TdsTypeNames = new ArrayList<String>();

			for (int i = 0; i < count1; i++) {
				String data = TCSTypeNameList.get(i).getText();
				TdsTypeNames.add(data);

				if (data.equalsIgnoreCase("Sale in cash of any goods (other than bullion/jewelry)")) {
					Thread.sleep(3000);
					TCSTypeChecklist1.get(i).click();
					break;
				}
				boolean tdschecklist1 = TCSTypeChecklist1.get(i).isSelected();
				System.err.println(tdschecklist1);
			}
			Thread.sleep(2000);
			getDriver().switchTo().parentFrame();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();

			String expUpdateMessage = "Updated Successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);

			Thread.sleep(3000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Accounts_CloseBtn));
			Accounts_CloseBtn.click();

			if (actUpdateMessage.equalsIgnoreCase(expUpdateMessage)) {
				return true;
			} else {
				return false;
			}

		}
	
		public static boolean checkSavingSalesInvoiceVocher1()
				throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {

			click(financialsMenu);
			click(financialsTransactionMenu);
			click(financialsTransactionsSalesMenu);
			click(SalesInvoiceVoucher);

			Thread.sleep(2000);

			click(newBtn);

			checkValidationMessage("Screen Opened");

			enterText(salesAccountTxt, "Sales - Computers");

			Thread.sleep(2000);

			tab(salesAccountTxt);

			enterText(CustomerAccounttxt, "Customer TCS");

			Thread.sleep(2000);
			tab(CustomerAccounttxt);

			click(salInvPlaceOfSupplyTxt);

			Thread.sleep(4000);

			tab(salInvPlaceOfSupplyTxt);

			click(select1stRow_1stColumn);

			enterText(enter_ItemTxt, "Item 1");

			Thread.sleep(2000);
			tab(enter_ItemTxt);

			enterText(enter_QuantityTxt, "10");

			tab(enter_QuantityTxt);

			enterText(enter_RateTxt, "2000");

			tab(enter_RateTxt);

			tab(enter_GrossTxt);

			click(select1stRow_6thColumn);

			enterText(enter_SalInvCGSTTxt, "9");

			tab(enter_SalInvCGSTTxt);

			enterText(enter_SalInvSGSTTxt, "9");

			tab(enter_SalInvSGSTTxt);

			tab(enter_SalInvIGSTTxt);

			tab(enter_SalInvCessTxt);

			tab(enter_SalInvTaxableValueTxt);

			Thread.sleep(2000);

			click(VoucherSaveBtn);

			Thread.sleep(5000);

			click(billRefPickIcon);

			Thread.sleep(3000);

			click(billRefOkBtn);

			System.out.println("************************checkSavingSalesInvoicVoucher**************************");

			String expMessage1 = "Voucher saved successfully";
			String expMessage2 = ": 1";
			String actMessage = checkValidationMessage(expMessage1);

			Thread.sleep(5000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(TCSFrame));
			getDriver().switchTo().frame(TCSFrame);

			Thread.sleep(5000);

			Select TCSType = new Select(TCSTypeDropdown);

			TCSType.selectByIndex(0);
			Thread.sleep(3000);

			String actTCSType = TCSType.getFirstSelectedOption().getText();
			String expTCSType = "Sale in cash of any goods (other than bullion/jewelry)";

			String actAmount = TCSDeductionPopupAmountTxt.getAttribute("value");
			String expAmount = "23600";

			String actTaxPercentage = TCSDeductionPopupTaxPercentageTxt.getAttribute("value");
			String expTaxPercentage = "7";

			String actTaxAmount = TCSDeductionPopupTaxAmountTxt.getAttribute("value");
			String expTaxAmount = "1652";

			String actTaxDeducted = TCSDeductionPopupTaxDeductedTxt.getAttribute("value");
			String expTaxDeducted = "1652";

			click(TCSPostBtn);

			getDriver().switchTo().parentFrame();

			System.out.println("*******************checkSaving Sales Invoce Voucher1******************");

			System.out.println("TCSType       : *" + actTCSType + "*  Value Expected  " + expTCSType);
			System.out.println("Amount        : *" + actAmount + "*  Value Expected  " + expAmount);
			System.out.println("TaxPercentage : *" + actTaxPercentage + "*  Value Expected  " + expTaxPercentage);
			System.out.println("TaxAmount     : *" + actTaxAmount + "*  Value Expected  " + expTaxAmount);
			System.out.println("TaxDeducted   : *" + actTaxDeducted + "*  Value Expected  " + expTaxDeducted);

			if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)
					&& actTCSType.equalsIgnoreCase(expTCSType) && actAmount.equalsIgnoreCase(expAmount)
					&& actTaxPercentage.equalsIgnoreCase(expTaxPercentage)
					&& actTaxAmount.equalsIgnoreCase(expTaxAmount) && actTaxDeducted.equalsIgnoreCase(expTaxDeducted)) {
				return true;
			} else {
				return false;
			}

		}

		public boolean checkValuesInSavedSalesVoucher1()
				throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
			Thread.sleep(5000);

			click(VoucherPreviousBtn);

			checkValidationMessage("Voucher Loaded Sucessfully");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccountTxt));
			String actsalesAccount = salesAccountTxt.getAttribute("data-focustext");
			String expsalesAccount = "Sales - Computers";

			System.out.println("=============" + actsalesAccount);

			System.out.println("=============" + expsalesAccount);

			Thread.sleep(5000);

			JavascriptExecutor js = (JavascriptExecutor) getDriver();
			js.executeScript("arguments[0].scrollIntoView();", tcsTxt);

			String actTcsBillAmt = tcsTxt.getAttribute("value");
			String expTcsBillAmt = "1,652.00";

			int count = voucherBodyRow1List.size();

			ArrayList<String> row1list = new ArrayList<String>();

			for (int i = 1; i < count; i++) {
				String data = voucherBodyRow1List.get(i).getText();
				row1list.add(data);
			}

			String actRow1List = row1list.toString();
			String expRow1List = "[Item 1, 10.00, 2,000.00, 20,000.00, 0.00, 9.00, 9.00, 0.00, 0.00, 20,000.00]";

			System.out.println("*****************checkValuesInSavedPaymentsVoucher1****************");

			System.out.println("actsalesAccount : " + actsalesAccount + "  Value Expected  " + expsalesAccount);
			System.out.println("actTcsBillAmt : " + actTcsBillAmt + "  Value Expected  " + expTcsBillAmt);
			System.out.println(actRow1List);
			System.out.println(expRow1List);

			if (actsalesAccount.equalsIgnoreCase(expsalesAccount) && actTcsBillAmt.equalsIgnoreCase(expTcsBillAmt)
					&& actRow1List.equalsIgnoreCase(expRow1List)) {
				return true;
			} else {
				return false;
			}
		}
	
	
	
	
	public static boolean clickOnVocherTypeAndvalidateTheVochersInTCSAfterSavingSalesInvoceVocher()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		Thread.sleep(3000);
		checkAvailablityOfTCSTab();

		Thread.sleep(2000);

		click(preferences_TCSVocherType);

		click(preferences_TcsVocherTypeRow1FieldsBtn);
		Thread.sleep(3000);

		if (preferences_TcsVocherTypeCGSTCheckBox.isSelected() == false) {

			click(preferences_TcsVocherTypeCGSTCheckBox);

		}

		if (preferences_TcsVocherTypeSGSTCheckBox.isSelected() == false) {

			click(preferences_TcsVocherTypeSGSTCheckBox);
		}

		if (preferences_TcsVocherTaxableCheckBox.isSelected() == false) {

			click(preferences_TcsVocherTaxableCheckBox);
		}

		Thread.sleep(3000);
		click(editScreenAndFooterOkBtn);

		click(tcsVocherTypeOkBtn);

		String expMessage1 = "Data Saved Successfully";
		String actMessage = checkValidationMessage(expMessage1);
		System.out.println("Actual Message     :  " + actMessage);
		System.out.println("exp Message        :  " + expMessage1);

		click(preferences_SaveBtn);
		
		getDriver().switchTo().parentFrame();


		if (actMessage.equalsIgnoreCase(expMessage1)) {
			return true;
		} else {
			return false;
		}

	}

	public static boolean checkSavingSalesInvoiceVocher2() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		
		click(financialsMenu);
		click(financialsTransactionMenu);
		click(financialsTransactionsSalesMenu);
		click(SalesInvoiceVoucher);

		Thread.sleep(2000);

		click(newBtn);

		checkValidationMessage("Screen Opened");

		enterText(salesAccountTxt, "Sales - Computers");

		Thread.sleep(2000);

		tab(salesAccountTxt);

		enterText(CustomerAccounttxt, "Customer TCS");

		Thread.sleep(2000);
		tab(CustomerAccounttxt);

		click(salInvPlaceOfSupplyTxt);

		Thread.sleep(4000);

		tab(salInvPlaceOfSupplyTxt);

		click(select1stRow_1stColumn);

		enterText(enter_ItemTxt, "Item 1");

		Thread.sleep(2000);
		tab(enter_ItemTxt);

		enterText(enter_QuantityTxt, "10");

		tab(enter_QuantityTxt);

		enterText(enter_RateTxt, "2000");

		tab(enter_RateTxt);

		tab(enter_GrossTxt);

		click(select1stRow_6thColumn);

		enterText(enter_SalInvCGSTTxt, "9");

		tab(enter_SalInvCGSTTxt);

		enterText(enter_SalInvSGSTTxt, "9");

		tab(enter_SalInvSGSTTxt);

		tab(enter_SalInvIGSTTxt);

		tab(enter_SalInvCessTxt);

		tab(enter_SalInvTaxableValueTxt);

		Thread.sleep(2000);

		click(VoucherSaveBtn);

		Thread.sleep(5000);

		click(billRefPickIcon);
		
		Thread.sleep(2000);

		click(billRefOkBtn);

		System.out.println(
				"************************checkSavingSalesInvoicVoucher2WithCustomerB2BTSForSuspendAndDelete**************************");

		String expMessage1 = "Voucher saved successfully";
		String expMessage2 = ": 2";
		String actMessage = checkValidationMessage(expMessage1);

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(TCSFrame));
		getDriver().switchTo().frame(TCSFrame);

		Thread.sleep(5000);

		Select TCSType = new Select(TCSTypeDropdown);
		TCSType.selectByIndex(0);
		Thread.sleep(3000);
		String actTCSType = TCSType.getFirstSelectedOption().getText();
		String expTCSType = "Sale in cash of any goods (other than bullion/jewelry)";

		String actAmount = TCSDeductionPopupAmountTxt.getAttribute("value");
		String expAmount = "20000";

		String actTaxPercentage = TCSDeductionPopupTaxPercentageTxt.getAttribute("value");
		String expTaxPercentage = "7";

		String actTaxAmount = TCSDeductionPopupTaxAmountTxt.getAttribute("value");
		String expTaxAmount = "1400";

		String actTaxDeducted = TCSDeductionPopupTaxDeductedTxt.getAttribute("value");
		String expTaxDeducted = "1400";

		click(TCSPostBtn);
		
		getDriver().switchTo().parentFrame();

		System.out.println("*******************checkSaving Sales Invoce Voucher1******************");

		System.out.println("TCSType       : " + actTCSType + "  Value Expected  " + expTCSType);
		System.out.println("Amount        : " + actAmount + "  Value Expected  " + expAmount);
		System.out.println("TaxPercentage : " + actTaxPercentage + "  Value Expected  " + expTaxPercentage);
		System.out.println("TaxAmount     : " + actTaxAmount + "  Value Expected  " + expTaxAmount);
		System.out.println("TaxDeducted   : " + actTaxDeducted + "  Value Expected  " + expTaxDeducted);

		if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)
				&& actTCSType.equalsIgnoreCase(expTCSType) && actAmount.equalsIgnoreCase(expAmount)
				&& actTaxPercentage.equalsIgnoreCase(expTaxPercentage) && actTaxAmount.equalsIgnoreCase(expTaxAmount)
				&& actTaxDeducted.equalsIgnoreCase(expTaxDeducted)) {
			return true;
		} else {
			return false;
		}

	}
	
	 public static boolean checkValuesInSavedSalesVoucher2() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
			Thread.sleep(3000);

			click(VoucherPreviousBtn);

			checkValidationMessage("Voucher Loaded Sucessfully");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccountTxt));
			String actsalesAccount = salesAccountTxt.getAttribute("data-focustext");
			String expsalesAccount = "Sales - Computers";

			JavascriptExecutor js = (JavascriptExecutor) getDriver();
			js.executeScript("arguments[0].scrollIntoView();", tcsTxt);

			String actTcsBillAmt = tcsTxt.getAttribute("value");
			String expTcsBillAmt = "1,400.00";

			int count = voucherBodyRow1List.size();

			ArrayList<String> row1list = new ArrayList<String>();

			for (int i = 1; i < count; i++) {
				String data = voucherBodyRow1List.get(i).getText();
				row1list.add(data);
			}

			String actRow1List = row1list.toString();
			String expRow1List = "[Item 1, 10.00, 2,000.00, 20,000.00, 0.00, 9.00, 9.00, 0.00, 0.00, 20,000.00]";

			System.out.println("*****************checkValuesInSavedPaymentsVoucher1****************");

			System.out.println("actsalesAccount : " + actsalesAccount + "  Value Expected  " + expsalesAccount);
			System.out.println("actTcsBillAmt : " + actTcsBillAmt + "  Value Expected  " + expTcsBillAmt);
			System.out.println(actRow1List);
			System.out.println(expRow1List);

			if (actsalesAccount.equalsIgnoreCase(expsalesAccount) && actTcsBillAmt.equalsIgnoreCase(expTcsBillAmt)
					&& actRow1List.equalsIgnoreCase(expRow1List)) {
				return true;
			} else {
				return false;
			}
		}
	 
	 public static boolean checkSavingSalesDebitNotes() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		 
			click(financialsMenu);
			click(financialsTransactionMenu);
			click(journalsMenu);
			click(salesDebitNotesMenu);
			click(new_newBtn);
			
			checkValidationMessage("Screen Opened");

			Thread.sleep(1000);
			enterText(vendorAccounttxt, "Customer tcs");

			Thread.sleep(2000);
			tab(vendorAccounttxt);
			click(select1stRow_1stColumn);
			enterText(enter_AccountTxt, "Bank Charges");

			Thread.sleep(2000);
			tab(enter_AccountTxt);

			enterText(enter_AmountTxt, "10");
			tab(enter_AmountTxt);

			Thread.sleep(1000);
			click(VoucherSaveBtn);

			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(MainFrame));

			getDriver().switchTo().frame(MainFrame);

			Thread.sleep(2000);

			click(postTax_InvoiceNoTxt);
			postTax_InvoiceNoTxt.sendKeys("3");
			postTax_InvoiceNoTxt.sendKeys(Keys.TAB);

			Thread.sleep(1000);
			click(postTax_PostBtn);

			Thread.sleep(2000);

			getDriver().switchTo().parentFrame();

			click(billRefPickIcon);
			Thread.sleep(2000);
			click(billRefOkBtn);

			String expMessage1 = "Voucher saved successfully";
			String expMessage2 = ": 1";
			String actMessage = checkValidationMessage(expMessage1);
			
			

			System.out.println("Message : " + actMessage + "  Value Expected  " + expMessage1 + " " + expMessage2);

			Thread.sleep(5000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(TCSFrame));
			getDriver().switchTo().frame(TCSFrame);

			Thread.sleep(5000);

			Select TCSType = new Select(TCSTypeDropdown);
			TCSType.selectByIndex(0);
			Thread.sleep(3000);
			String actTCSType = TCSType.getFirstSelectedOption().getText();
			String expTCSType = "Sale in cash of any goods (other than bullion/jewelry)";

			String actAmount = TCSDeductionPopupAmountTxt.getAttribute("value");
			String expAmount = "10";

			String actTaxPercentage = TCSDeductionPopupTaxPercentageTxt.getAttribute("value");
			String expTaxPercentage = "7";

			String actTaxAmount = TCSDeductionPopupTaxAmountTxt.getAttribute("value");
			String expTaxAmount = "1";

			String actTaxDeducted = TCSDeductionPopupTaxDeductedTxt.getAttribute("value");
			String expTaxDeducted = "1";

			click(TCSPostBtn);
			
			getDriver().switchTo().parentFrame();

			System.out.println("*******************checkSaving Sales Invoce Voucher1******************");

			System.out.println("TCSType       : " + actTCSType + "  Value Expected  " + expTCSType);
			System.out.println("Amount        : " + actAmount + "  Value Expected  " + expAmount);
			System.out.println("TaxPercentage : " + actTaxPercentage + "  Value Expected  " + expTaxPercentage);
			System.out.println("TaxAmount     : " + actTaxAmount + "  Value Expected  " + expTaxAmount);
			System.out.println("TaxDeducted   : " + actTaxDeducted + "  Value Expected  " + expTaxDeducted);

			if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)
					&& actTCSType.equalsIgnoreCase(expTCSType) && actAmount.equalsIgnoreCase(expAmount)
					&& actTaxPercentage.equalsIgnoreCase(expTaxPercentage) && actTaxAmount.equalsIgnoreCase(expTaxAmount)
					&& actTaxDeducted.equalsIgnoreCase(expTaxDeducted)) {
				return true;
			} else {
				return false;
			}

		}
	 
		public static boolean checkSavingReceipts1()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

			click(financialsMenu);
			click(financialsTransactionMenu);
			click(cashAndBankMenu);
			click(receiptsMenu);
			click(new_newBtn);
			checkValidationMessage("Screen Opened");

			Thread.sleep(2000);

			enterText(cashAnsBankActTxt, "Cash");
			Thread.sleep(2000);

			tab(cashAnsBankActTxt);

			click(select1stRow_1stColumn);
			enterText(enter_AccountTxt, "Customer tcs");
			Thread.sleep(2000);
			tab(enter_AccountTxt);

			Thread.sleep(2000);
			enterText(enter_AmountTxt, "100");

			Thread.sleep(2000);
			tab(enter_AmountTxt);

			click(billRefPickIcon);
			Thread.sleep(2000);
			click(billRefOkBtn);

			click(VoucherSaveBtn);

			String expMessage1 = "Voucher saved successfully";
			String expMessage2 = ": 1";
			String actMessage = checkValidationMessage(expMessage1);

			System.out.println("Message : " + actMessage + "  Value Expected  " + expMessage1 + " " + expMessage2);

			Thread.sleep(5000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(TCSFrame));
			getDriver().switchTo().frame(TCSFrame);

			Thread.sleep(5000);

			Select TCSType = new Select(TCSTypeDropdown);
			TCSType.selectByIndex(0);
			Thread.sleep(3000);
			String actTCSType = TCSType.getFirstSelectedOption().getText();
			String expTCSType = "Sale in cash of any goods (other than bullion/jewelry)";

			String actAmount = TCSDeductionPopupAmountTxt.getAttribute("value");
			String expAmount = "100";

			String actTaxPercentage = TCSDeductionPopupTaxPercentageTxt.getAttribute("value");
			String expTaxPercentage = "7";

			String actTaxAmount = TCSDeductionPopupTaxAmountTxt.getAttribute("value");
			String expTaxAmount = "7";

			String actTaxDeducted = TCSDeductionPopupTaxDeductedTxt.getAttribute("value");
			String expTaxDeducted = "7";

			click(TCSPostBtn);

			System.out.println("*******************checkSaving Sales Invoce Voucher1******************");

			System.out.println("TCSType       : " + actTCSType + "  Value Expected  " + expTCSType);
			System.out.println("Amount        : " + actAmount + "  Value Expected  " + expAmount);
			System.out.println("TaxPercentage : " + actTaxPercentage + "  Value Expected  " + expTaxPercentage);
			System.out.println("TaxAmount     : " + actTaxAmount + "  Value Expected  " + expTaxAmount);
			System.out.println("TaxDeducted   : " + actTaxDeducted + "  Value Expected  " + expTaxDeducted);

			getWaitForAlert();
			String actAlert = getAlert().getText();
			String expAlert = "TCS Journal entry(JV:1) successfully posted.";

			getAlert().accept();

			getDriver().switchTo().parentFrame();

			if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)
					&& actTCSType.equalsIgnoreCase(expTCSType) && actAmount.equalsIgnoreCase(expAmount)
					&& actTaxPercentage.equalsIgnoreCase(expTaxPercentage)
					&& actTaxAmount.equalsIgnoreCase(expTaxAmount) && actTaxDeducted.equalsIgnoreCase(expTaxDeducted)) {
				return true;
			} else {
				return false;
			}

		}
		
		public static boolean validateVocherInTCSJvAfterSavingReceipts1() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
			
			click(financialsMenu);
			click(financialsTransactionMenu);
			click(journalsMenu);
			click(tcsJVMenu);

			Thread.sleep(3000);

			int count = grid_VoucherNoList.size();

			for (int i = 0; i < count; i++) {
				String VoucherNo = grid_VoucherNoList.get(i).getText();

				if (VoucherNo.equalsIgnoreCase("1")) {
					if (grid_CheckBoxList.get(i).isSelected() == false) {
						getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
						break;
					}
				}
			}
			
			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNoTxt));
			String actDocumentNo = documentNoTxt.getAttribute("value");
			String expDocumentNo = "1";
			System.out.println("actDocumentNo           :"+actDocumentNo);
			
			Thread.sleep(2000);

			//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tcsVocherno));
			String actlTcsVocherno = tcsVocherno.getAttribute("value");
			String expTcsVocherno = "Rct:1";

			int count2 = jsVocherBodyRow1List.size();

			ArrayList<String> row1list = new ArrayList<String>();

			for (int i = 1; i < count2; i++) {
				String data = jsVocherBodyRow1List.get(i).getText();
				row1list.add(data);
			}

			String actRow1List = row1list.toString();
			String expRow1List = "[Customer TCS, Authorized Capital, 7.00, New Reference]";
			
			 System.out.println("*****************checkValuesSavedJournalEntriesVoucher1****************");
			 
			 System.out.println("actDocumentNo : "+actDocumentNo+"  Value Expected  "+expDocumentNo);
			 System.out.println("TCSJVNo : "+actlTcsVocherno+"  Value Expected  "+expTcsVocherno);
			 System.out.println(actRow1List);
			 System.out.println(expRow1List);
			 Thread.sleep(2000);
				if (actRow1List.equalsIgnoreCase(expRow1List) && actDocumentNo.equalsIgnoreCase(expDocumentNo)
						&& (actlTcsVocherno.equalsIgnoreCase(expTcsVocherno))) {
					
					click(tcsJvCloseBtn);

					return true;
				} else {
					
					click(tcsJvCloseBtn);
					
					return false;
				}

		}
		
		public static boolean checkSavingReceipts2() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
			Thread.sleep(3000);
			click(financialsMenu);
			click(financialsTransactionMenu);
			click(cashAndBankMenu);
			click(receiptsMenu);
			Thread.sleep(3000);
			click(newBtn);
			checkValidationMessage("Screen Opened");

			Thread.sleep(2000);

			enterText(cashAnsBankActTxt, "Cash");
			Thread.sleep(2000);

			tab(cashAnsBankActTxt);

			click(select1stRow_1stColumn);
			enterText(enter_AccountTxt, "Customer tcs");
			Thread.sleep(2000);
			tab(enter_AccountTxt);

			Thread.sleep(2000);
			enterText(enter_AmountTxt, "100");

			Thread.sleep(2000);
			tab(enter_AmountTxt);
			
			Thread.sleep(2000);
			click(referenceGridCheckBox);
			Thread.sleep(1000);
			click(billRefPickIcon);
			Thread.sleep(2000);
			click(billRefOkBtn);

			click(VoucherSaveBtn);

			Thread.sleep(2000);

			String expMessage1 = "Voucher saved successfully";
			String expMessage2 = ": 2";
			String actMessage = checkValidationMessage(expMessage1);

			getWaitForAlert();
			String actAlert = getAlert().getText();
			String expAlert = "TCS already got deducted for the adjusted document's. So no TCS will be collected for this document.";

			getAlert().accept();
			
			System.out.println("Message : " + actMessage + "  Value Expected  " + expMessage1 + " " + expMessage2);
			System.out.println("actAlert          :"+actAlert);
			System.out.println("expAlert          :"+expAlert);
			if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)
					&& actAlert.equalsIgnoreCase(expAlert)) {
				return true;
			} else {
				return false;
			}

		}
		
		public static String checkDownloadedFileName(WebDriver driver) throws InterruptedException
		{
			String mainWindow = driver.getWindowHandle();

			JavascriptExecutor js = (JavascriptExecutor)driver;
			js.executeScript("window.open()");

			for(String winHandle : driver.getWindowHandles())
			{
				driver.switchTo().window(winHandle);
			}

			driver.get("chrome://downloads");

			JavascriptExecutor js1 = (JavascriptExecutor)driver;

			String fileName = (String) js1.executeScript("return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content #file-link').text");

			System.err.println("Download deatils");
			System.out.println("File Name :-" + fileName);

			Thread.sleep(2000);
			driver.close();

			driver.switchTo().window(mainWindow);

			return fileName;
		}


	 
		
		public static boolean checkValidateExportingTCSTransactionDetails() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		        Thread.sleep(2000);
		        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionDetailsExportBtn));
		        transactionDetailsExportBtn.click();
		        
		        String actconfirmTxt = ExcelConfirmTxt.getText();
		        String expConfirmTxt = "Do you want to export report to the Excel?";
		        
		        System.out.println("actconfirmTxt: "+actconfirmTxt);
		        System.out.println("expConfirmTxt: "+expConfirmTxt);
		        
		        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExcelConfirmYesBtn));
		        ExcelConfirmYesBtn.click();
		        
		        actFileName = checkDownloadedFileName(getDriver());
		        String expFileName = "Transaction Details";
		        
		        System.out.println("File Name : "+actFileName+"  Value Expected  "+expFileName);
		        
		        String actExcelfile = getBaseDir()+"\\autoIt\\ExportFiles\\"+actFileName;
		        
		        ExcelReader excelReader = new ExcelReader(actExcelfile);
		        String sheet = "Sheet1";
		        
		        
		        if (actFileName.startsWith(expFileName)) 
		        {
		                return true;
		        } 
		        else 
		        {
		                return false;
		        }

		}
		
		

public static boolean checkNValidateExportedTCSTransactionDetailsExcel() throws EncryptedDocumentException, InvalidFormatException, IOException
{

        ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());
        
        String actExcelfile = getBaseDir()+"\\autoIt\\ExportFiles\\"+actFileName;
        String expExcelfile = getBaseDir()+"\\autoIt\\ImportFiles\\ExportTCSTransactionDetails.xlsx";
        String sheet = "Sheet1";
           
        Calendar cal=Calendar.getInstance();
        SimpleDateFormat currentDate = new SimpleDateFormat("dd-MM-yyyy");
        String presentDate = currentDate.format(cal.getTime());
        
        System.err.println("Present date : "+presentDate);
        
        
        excelReader.setCellData(expExcelfile, sheet, 5, 0, presentDate);
        excelReader.setCellData(expExcelfile, sheet, 6, 0, presentDate);
        excelReader.setCellData(expExcelfile, sheet, 7, 0, presentDate);
        excelReader.setCellData(expExcelfile, sheet, 8, 0, presentDate);
       
        excelReader.setCellData(expExcelfile, sheet, 5, 20, presentDate);
        excelReader.setCellData(expExcelfile, sheet, 6, 20, presentDate);
        excelReader.setCellData(expExcelfile, sheet, 7, 20, presentDate);
        excelReader.setCellData(expExcelfile, sheet, 8, 20, presentDate);
        

        FileInputStream fip1 = new FileInputStream(actExcelfile);
        Workbook workbook1  = WorkbookFactory.create(fip1);
        
        FileInputStream fip2 = new FileInputStream(expExcelfile);
        Workbook workbook2  = WorkbookFactory.create(fip2);
        
        boolean result = excelReader.verifyDataInExcelBookAllSheets(workbook1, workbook2);
        
        System.err.println(result);
        
        if (result) 
        {
                return true;
        }
        else 
        {
                return false;
        }

}
@FindBy(xpath = "//div[@id='btnsaveChallan']")
private static WebElement tcsChallanSaveBtn;
public static boolean checkOPenTDSChallan() throws InterruptedException
{
	Thread.sleep(3000);
 	getDriver().navigate().refresh();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(statutoryMenu));
	statutoryMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tdsTcsMenu));
	tdsTcsMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tcsChallanMenu));
	tcsChallanMenu.click();
	Thread.sleep(3000);
	
	getDriver().switchTo().frame(iframeElement);
	Thread.sleep(3000);
	
	boolean actTdsChallanTitleTxt = tcsChallanSaveBtn.isDisplayed();
	boolean expTdsChallanTitleTxt = true;
	
	if (actTdsChallanTitleTxt== expTdsChallanTitleTxt)
	{
		return true;
	}
	else
	{
		return false;
	}

}

public static boolean checkGenerateChallanForAllVouchersInTDStransactionDetails() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
		//TCSTransactionDetailsPage.checkOpenTCSTransactionDetails();
        
        click(transactionDetailsStAllYesOrNo);
        
        Thread.sleep(2000);
        click(transactionDetailsGenerateChallan);
        
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageTxt));
        String actMessage1 = errorMessageTxt.getText();
        String expMessage = "Data Saved Successfully";
        
        click(errorMessageClose);
        Thread.sleep(3000);
        
        click(tcsChallnCloseBtn);
        
        
        getDriver().switchTo().parentFrame();
        
        getDriver().navigate().refresh();
        Thread.sleep(3000);
        checkOPenTDSChallan();
        
        Thread.sleep(3000);
        
        
        Thread.sleep(2000);
        Select serialno = new Select(tcsChallanSerialNoDropDown);
        serialno.selectByVisibleText("2");
        
        String actTdsSectionType=tcsChallanTcsTypetxt.getText();
        String expTdsSectionType="Payment to Contractor / Subcontractor / Advertisements ,  Insurance commission";
        
        System.out.println("actTdsSectionType:"+actTdsSectionType);
        System.out.println("expTdsSectionType:"+expTdsSectionType);
        
        Select challnTYpe = new Select(tcsChallanChallanTypeDropDown);
        challnTYpe.selectByVisibleText("Company");
        
        enterText(tcsChallanChallanNoTxt,"4876");
        
        String actCurrentdate=tcsChallanDateOfChallanTxt.getText();
        String expCurrentdate = getCurrentDateF2();
        
        System.out.println("actCurrentdate:"+actCurrentdate);
        System.out.println("expCurrentdate:"+expCurrentdate);
        
        Select month = new Select(tcsChallanMonthDropDown);
        month.selectByVisibleText(getMonth());
        
        Select bank = new Select(tcsChallanBankNameDropDown);
        bank.selectByVisibleText("Bank");
        
        Select payment = new Select(tcsChallanPaymentModeDropDown);
        payment.selectByVisibleText("Cash");
        
        
        click(tcsChallanSaveBtn);
        
        String expSaveMessage = "Challan Saved Successfully";
        
        String actSaveMessage = checkValidationMessage(expSaveMessage);
        
        if (actSaveMessage.equalsIgnoreCase(expSaveMessage) && actTdsSectionType.equalsIgnoreCase(expTdsSectionType) && actCurrentdate.equalsIgnoreCase(expCurrentdate)) 
        {
                return true;
        }
        else 
        {
                return false;
        }

}


	}
