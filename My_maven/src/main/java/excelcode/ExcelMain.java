package excelcode;

import java.io.IOException;

public class ExcelMain 
{
	public static void main(String args[]) throws IOException
	{
		Excel ob=new Excel();
		String a= ob.readData(0,0);
		System.out.println("Value of a is"+a);
		String b= ob.readData(0,1);	
		System.out.println("Value of b is"+b);
	}
}
