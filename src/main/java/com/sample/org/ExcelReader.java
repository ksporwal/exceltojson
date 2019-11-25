package com.sample.org;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import com.opencsv.CSVWriter;

public class ExcelReader {
	public static final String SAMPLE_XLSX_FILE_PATH = "C:\\Users\\porwal\\Downloads\\MOD_IND_21NOV2019.xlsx";
	public static final String CSV_OUTPUT_FILE_PATH = "C:\\Users\\porwal\\Downloads\\modoutput.csv";
	@SuppressWarnings("unchecked")
	public static void main(String[] args) throws IOException, InvalidFormatException {
		InputStream is = new FileInputStream(SAMPLE_XLSX_FILE_PATH);
		// Creating a Workbook from an Excel file (.xls or .xlsx)
		Workbook workbook = WorkbookFactory.create(is);


		// Retrieving the number of sheets in the Workbook
		System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");


		// Getting the Sheet at index zero
		Sheet sheet = workbook.getSheetAt(0);
	
		
		// 1. You can obtain a rowIterator and columnIterator and iterate over them
		System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
		Iterator<Row> rowIterator = sheet.rowIterator();
		int i=1;
		 CSVWriter csvWriter = new CSVWriter(new FileWriter(CSV_OUTPUT_FILE_PATH));
		 csvWriter.writeNext(new String[]{"slno", "Name","Metadata","Lattitude","Longitude"});
	if(rowIterator.hasNext())
		rowIterator.next();
		 while (rowIterator.hasNext()) {
			
			Row row = rowIterator.next();
			JSONObject obj = new JSONObject();
			JSONObject obj1 = new JSONObject();
			JSONObject obj2 = new JSONObject();
			JSONObject obj3 = new JSONObject();
			JSONObject obj4 = new JSONObject();
			JSONObject obj5 = new JSONObject();
			JSONObject obj6 = new JSONObject();
			JSONObject obj7 = new JSONObject();
			
		
			JSONArray arr = new JSONArray();
			
		
			String Name = row.getCell(5).toString();
		
			double pcode = Double.parseDouble(row.getCell(52).toString());
			int pcode1 = (int)pcode;
			String pcode2 = String.valueOf(pcode1);
			
			String Latitude = row.getCell(21).toString();
			String Longitude = row.getCell(22).toString();
			
			String str = row.getCell(6).getStringCellValue();
		
			
			
			obj.put("question", "Categories");
			obj.put("answer", "POINT OFINTEREST");
			arr.add(obj);

			obj1.put("question", "Sub Categories");
			obj1.put("answer",  str );
			arr.add(obj1);
			
			obj2.put("question", "Features");
			obj2.put("answer", row.getCell(7) +"");
			arr.add(obj2);
			
			obj3.put("question", "NAME");
			obj3.put("answer", row.getCell(5)+"");
			arr.add(obj3);

			obj4.put("question", "LOCATION NAME");
			obj4.put("answer", row.getCell(50)+"");
			arr.add(obj4);
			
			
			obj5.put("question", "TYPE");
			obj5.put("answer", str);
			arr.add(obj5);
			
			
			obj6.put("question", "ADDITIONAL INFORMATION");
			obj6.put("answer", row.getCell(48) + ","+pcode2);
			arr.add(obj6);
			
			obj7.put("question", "SURVEY NO");
			obj7.put("answer", row.getCell(44)+"");
			arr.add(obj7);
	
			
			String data = arr.toJSONString();	
			String data1 = data.replaceAll("null", "");		//For null values
			String newstr = data1.replace("\\/", "/");		//For backslash replacing
			
	//		System.out.println(obj5);
			//ADD DATA TO CSV
			csvWriter.writeNext(new String[] {String.valueOf(i), Name,newstr,Latitude,Longitude});
		
			i++;
	
		}
		
		
		// Closing the workbook
		workbook.close();
		csvWriter.close();
		System.out.println("\nworkbook closed");
	}
}