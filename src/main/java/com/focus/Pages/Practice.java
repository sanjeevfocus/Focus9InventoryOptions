package com.focus.Pages;

import java.io.BufferedReader;
import java.io.DataInputStream;

import java.io.FileInputStream;

import java.io.InputStreamReader;

public class Practice {

	public static String get() {
		String data = "";
		try {

			String name = "";
			FileInputStream fstream = new FileInputStream(".\\test.txt");
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			while ((name = br.readLine()) != null) {
				data += name + "\n";

			}

			in.close();
		} catch (Exception e) {
			System.err.println("Error: " + e.getMessage());
		}
		return data;
	}

	public static void main(String[] args) throws Exception {

		Practice p = new Practice();
		String str = get();
		System.err.println(str);

	}

}
