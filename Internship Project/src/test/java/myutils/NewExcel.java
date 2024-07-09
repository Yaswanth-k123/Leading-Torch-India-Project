package myutils;
import java.util.*;

import java.util.Arrays;
import java.util.Date;
import java.util.concurrent.TimeUnit;
import java.text.SimpleDateFormat;
import java.io.*;
import java.util.Comparator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class NewExcel {
	//this method returns list of column values for the specified row
	public static ArrayList<String> getCellData(int j,XSSFSheet sheet) throws Exception 
	{

		 ArrayList<String> list=new ArrayList<String>();
		 String citem,name;
		 Date date1,date2;
			citem=sheet.getRow(j).getCell(0).getStringCellValue();
			name=sheet.getRow(j).getCell(3).getStringCellValue();
			 date1=sheet.getRow(j).getCell(1).getDateCellValue();
			 date2=sheet.getRow(j).getCell(2).getDateCellValue();
			 Calendar calendar = Calendar.getInstance();
	         calendar.setTime(date1);
	         int year = calendar.get(Calendar.YEAR);
	         int month = calendar.get(Calendar.MONTH) + 1; // Note: Calendar.MONTH is zero-based
	         int day = calendar.get(Calendar.DAY_OF_MONTH);
	         // Format the date as desired (e.g., "MM/dd/yyyy")
	         String d1 = String.format("%02d/%02d/%04d", day,month, year);
			
	    	 calendar = Calendar.getInstance();
	         calendar.setTime(date2);
	         year = calendar.get(Calendar.YEAR);
	         month = calendar.get(Calendar.MONTH) + 1; // Note: Calendar.MONTH is zero-based
	         day = calendar.get(Calendar.DAY_OF_MONTH);


	         String d2 = String.format("%02d/%02d/%04d", day,month, year);
	    
	    	list.add(citem);
	    	list.add(d1);
	    	list.add(d2);
	    	list.add(name);
	    	return list;
		}
	
	//this method sets the heading to the specified sheet
		public static void setHeading(XSSFSheet sheet1) {
			XSSFSheet sheet=sheet1;
			Row r1;
			Cell c1;
			 r1=sheet.createRow(0);
		     
			   c1=r1.createCell(0);
			   c1.setCellValue("Configuration item");
			   c1=r1.createCell(1);
			   c1.setCellValue("1TS Created Date");
			   c1=r1.createCell(2);
			   c1.setCellValue("1TS Rdesolved");
			   c1=r1.createCell(3);
			   c1.setCellValue("Assigned to");
			   c1=r1.createCell(4);
			   c1.setCellValue("Days");
		}
	public static void main(String args[]) throws Exception
	{
		
		
			String inputExcelPath="./newdata/data.xlsx";                   //input excel file
			XSSFWorkbook workbook1=new XSSFWorkbook(inputExcelPath);
			XSSFSheet input=workbook1.getSheet("Sheet1");                  //getting reference of input sheet
			String outputExcelPath="./newdata/output.xlsx";                //output excel file
			XSSFWorkbook workbook=new XSSFWorkbook();                      //create an excel workbook
			HashMap<String ,Integer> map=new HashMap<String, Integer>();   // creating object of HashMap
			 ArrayList<String> list=new ArrayList<String>();               // creating object of ArrayList     
			 
			  // declaring variables
			    int temp,i=0,n=0,m=0;
				Row r1;
				Cell c1;
				String citem = null,date1 = null,date2 = null,name = null,s1,s2,firstCell;
				long days=0,diff=0;
				 SimpleDateFormat myFormat=new SimpleDateFormat("dd/MM/yyyy");
			
				 int rowcount=input.getPhysicalNumberOfRows(); 
			for(i=1;i<rowcount;i++) {
				citem=input.getRow(i).getCell(0).getStringCellValue();
				if(map.containsKey(citem)) {
					   temp=map.get(citem);
					   temp++;
					   map.put(citem,temp);
					   
				}
				else {
	               map.put(citem, 1);
				}
			}
			System.out.println(map);
		    final int columnToSort = 3;                                     // Sort based on the second column (index 1)
			for(String s:map.keySet()) {

				XSSFSheet sheet=workbook.createSheet();                     //creating sheet in the output excel file
				String arr[][]=new String[map.get(s)][5];                   // declaring an array
	            n=0;
	            m=0;
	            setHeading(sheet);            // this method sets the heading to the specified sheet
				for(i=1;i<rowcount;i++) {
					firstCell=input.getRow(i).getCell(0).getStringCellValue();
					if(s.equals(firstCell)) {
						list=getCellData(i,input);                      
						
						       citem=list.get(0);			                         // fetching the values
							   date1=list.get(1);
							   date2=list.get(2);
							   name=list.get(3);
							   s1=date1;
							   s2=date2;
							   
							   try {
								   Date d1=myFormat.parse(s1);
								   Date d2=myFormat.parse(s2);
								   diff=d2.getTime()-d1.getTime();
								    days= TimeUnit.DAYS.convert(diff,TimeUnit.MILLISECONDS);       // calculating difference bw days
								   
							   }
							   catch(Exception e) {
								   System.out.println(e);
							   }
					
							   arr[n][m]=citem;                               // assigning the values into array
							   m++;
							   arr[n][m]=date1;
							   m++;
							   arr[n][m]=date2;
							   m++;
							   arr[n][m]=name;
							   m++;
							   arr[n][m]=days+"";
								  m=0;
								  n++;   
						     }
					
					else 
					{
						continue;
					}
					}
			  
 
			       Arrays.sort(arr, new Comparator<String[]>() {                                //sorting the array
			           public int compare(String[] array1, String[] array2) {
			               return array1[columnToSort].compareTo(array2[columnToSort]);
			           }
			       }); 
		
			   	for(n=0;n<arr.length;n++) {                         // insering data into  output excel shhet
					   r1=sheet.createRow(n+1);                     // creating row
					   c1=r1.createCell(0);                         //creating cell
					   c1.setCellValue(arr[n][0]);
					   c1=r1.createCell(1);
					   c1.setCellValue(arr[n][1]);
					   c1=r1.createCell(2);
					   c1.setCellValue(arr[n][2]);
					   c1=r1.createCell(3);
					   c1.setCellValue(arr[n][3]);
					   c1=r1.createCell(4);
					   c1.setCellValue(arr[n][4]);
				}
				
				}
	     
	  

		
			
			
			  try {
				    FileOutputStream out=new FileOutputStream(outputExcelPath);
				    workbook.write(out);
				    out.close();			    workbook.close();
				    System.out.println("inserted successfully");
				    }
				    catch(Exception e) {
				    	System.out.println(e);	    
					}	
			 

	    }

	}

    
    
    
