package com.focus.Pages;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;

public class Practice2 {
	
	public static String m1() throws IOException {

		File file = new File("C:\\Users\\Roja\\Desktop\\hi.txt");

		BufferedReader br = new BufferedReader(new FileReader(file));

		

		String st;
		while ((st = br.readLine()) != null)
			System.out.println(st);
		return st;
		
	}

	public static void main(String[] args) throws Exception {

		Practice2 p1=new Practice2();
		String str=m1();
		System.err.println(str);
	}

}
