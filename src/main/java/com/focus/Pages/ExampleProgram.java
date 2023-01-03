package com.focus.Pages;

public class ExampleProgram {
	
	String a;
	
	 public void m1(String a)
	 
	  {
		 this.a=a;		
		 System.out.println(a);
		 
	  }
	 
	 public void m2(String a1)
	  {
		 System.out.println(a);
		 System.out.println(a1);
	  }
	 
	 


	public static void main(String[] args) 
	
	{
		ExampleProgram a=new ExampleProgram();
		a.m1("sanjeev");
		a.m2("Raj");
		
	}

}
