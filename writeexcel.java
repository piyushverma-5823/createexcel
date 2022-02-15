package test;
import  java.io.*;  
import  org.apache.poi.hssf.usermodel.HSSFSheet;  
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import  org.apache.poi.hssf.usermodel.HSSFRow;  
public class writeexcel  
{  
public static void main(String[]args)   
{  
	//creation of excel file
	try   
	{  
	String filename = "C:\\Users\\piyush.verma\\Desktop\\StudentsDetail.xlsx";  
	FileOutputStream fileOut = new FileOutputStream(filename);  
	fileOut.close();  
	System.out.println("Excel file has been Created successfully.");  
	}   
	catch (Exception e)   
	{  
	e.printStackTrace();  
	} 
	//Writing the excel file
try   
{  
//declare file name to be create   
String filename = "C:\\Users\\piyush.verma\\Desktop\\StudentsDetail.xlsx";   
//creating an instance of HSSFWorkbook class  
HSSFWorkbook workbook = new HSSFWorkbook();  
//invoking creatSheet() method and passing the name of the sheet to be created   
HSSFSheet sheet = workbook.createSheet("Basic Details");   
//creating the 0th row using the createRow() method  
HSSFRow rowhead = sheet.createRow((short)0);  
//creating cell by using the createCell() method and setting the values to the cell by using the setCellValue() method  
rowhead.createCell(0).setCellValue("S.No.");  
rowhead.createCell(1).setCellValue("Student Name");  
rowhead.createCell(2).setCellValue("Roll Number");  
rowhead.createCell(3).setCellValue("e-mail");  
rowhead.createCell(4).setCellValue("Current Percentage");  
//creating the 1st row  
HSSFRow row = sheet.createRow((short)1);  
//inserting data in the first row  
row.createCell(0).setCellValue("1");  
row.createCell(1).setCellValue("Piyush Verma");  
row.createCell(2).setCellValue("1814310139");  
row.createCell(3).setCellValue("piyush.verma@qualitestgroup.com");  
row.createCell(4).setCellValue("82.00");  
//creating the 2nd row  
HSSFRow row1 = sheet.createRow((short)2);  
//inserting data in the second row  
row1.createCell(0).setCellValue("2");  
row1.createCell(1).setCellValue("Pradumn Gaud");  
row1.createCell(2).setCellValue("1814310140");  
row1.createCell(3).setCellValue("pradumn.gaud@qualitestgroup.com"); 
row1.createCell(4).setCellValue("85.00"); 
//creating the 3rd row  
HSSFRow row2 = sheet.createRow((short)3);  
//inserting data in the third row  
row2.createCell(0).setCellValue("3");  
row2.createCell(1).setCellValue("Prashant Saxena");  
row2.createCell(2).setCellValue("1814310147");  
row2.createCell(3).setCellValue("prashant.saxena@qualitestgroup.com");  
row2.createCell(4).setCellValue("80.00");
//creating the 4th row  
HSSFRow row3 = sheet.createRow((short)4);  
//inserting data in the third row  
row3.createCell(0).setCellValue("4");  
row3.createCell(1).setCellValue("Apoorva Rauniyar");  
row3.createCell(2).setCellValue("1814310141");  
row3.createCell(3).setCellValue("apoorva.rauniyar@qualitestgroup.com");  
row3.createCell(4).setCellValue("83.00");
FileOutputStream fileOut = new FileOutputStream(filename);  
workbook.write(fileOut);  
//closing the Stream  
fileOut.close();  
//closing the workbook  
workbook.close();  
//prints the message on the console  
System.out.println("Excel file has been written successfully.");  
}   
catch (Exception e)   
{  
e.printStackTrace();  
}  
}  
}  