package com.focus.base;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

public class ActionEngine extends BaseEngine 
{
static private WebElement element;
static private List<WebElement> elements;


    static public WebElement identifyElement(String locMech,String locType) 
    {
	switch(locType)
	{
	case "id": element = getDriver().findElement(By.id(locMech));
	if (element.isDisplayed()&&element.isEnabled()) {
		System.out.println("It is enabled");
	}
	else
	{
		System.out.println("Disabled");
	}
	break;
	
	case "name":element = getDriver().findElement(By.name(locMech));
	if (element.isDisplayed()&&element.isEnabled()) {
		System.out.println("It is enabled");
	}
	else{
		System.out.println("Disabled");
	}
	break;
	case "class":element = getDriver().findElement(By.className(locMech));
	if (element.isDisplayed()&&element.isEnabled()) {
		System.out.println("It is enabled");
	}
	else{
		System.out.println("Disabled");
	}
	break;
	case "xpath":element = getDriver().findElement(By.xpath(locMech));
	if (element.isDisplayed()&&element.isEnabled()) {
		System.out.println("It is enabled");
	}
	else{
		System.out.println("Disabled");
	}
	break;
	case "css":element = getDriver().findElement(By.cssSelector(locMech));
	if (element.isDisplayed()&&element.isEnabled()) {
		System.out.println("It is enabled");
	}
	else{
		System.out.println("Disabled");
	}
	break;
	case "tagName":element = getDriver().findElement(By.tagName(locMech));
	if (element.isDisplayed()&&element.isEnabled()) {
		System.out.println("It is enabled");
	}
	else{
		System.out.println("Disabled");
	}	
	break;
	case "linkText":element = getDriver().findElement(By.linkText(locMech));
	if (element.isDisplayed()&&element.isEnabled()) {
		System.out.println("It is enabled");
	}
	else{
		System.out.println("Disabled");
	}
	break;
	case "partialLinkText":element = getDriver().findElement(By.partialLinkText(locMech));
	if (element.isDisplayed()&&element.isEnabled()) {
		System.out.println("It is enabled");
	}
	else{
		System.out.println("Disabled");
	}
	break;
	}
		return element;
	}
    
    
	public static void enterUrl(String url) 
	{
		getDriver().get(url);
	}
	
	

    
}

