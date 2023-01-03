package com.focus.Pages;

public class Practice3 {
	
	int rollno;
	String name;
	float fee;

	public void Student(int rollno, String name, float fee) {
		this.rollno = rollno;
		this.name = name;
		this.fee = fee;
	}

	void display() {
		System.out.println(rollno + " " + name + " " + fee);

	}
	
	
	public static void main(String[] args) {
		
		Practice3 p=new Practice3();
		p.display(); 
		
	}
}
