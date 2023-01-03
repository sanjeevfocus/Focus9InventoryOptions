package com.focus.Pages;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class CheckDate {

	public static String getDate() {
		
		
		DateFormat dateFormat = new SimpleDateFormat("MMMM yyyy");
		Calendar cal = Calendar.getInstance();
		//System.out.println(dateFormat.format(cal.getTime()));
		
		return dateFormat.format(cal.getTime());
		
	}

}
