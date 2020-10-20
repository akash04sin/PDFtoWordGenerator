package com.app;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.app.pojo.ExcelSheetDetails;
import com.app.pojo.FlawDetail;
import com.app.pojo.PDFFileDetails;

public class ReadExcel {
	
	public static void readExcel(String filePath, PDFFileDetails pdfFileDetails) throws Exception{
		
		
		int sheetIndex = 0;
		try {
			//Create a workbook
			Workbook workbook = WorkbookFactory.create(new File(filePath));
					
			//Get the appropraite sheet
			String eaiCode = pdfFileDetails.getEaicode();
			sheetIndex = getSheetIndex(workbook,eaiCode); 
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			generateAppInfo(sheet,pdfFileDetails);
			
			workbook.close();
			
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static int getSheetIndex(Workbook workbook, String eaiCode){
		
		int index = 0;
		Iterator<Sheet> sheetItr = workbook.sheetIterator();
		while(sheetItr.hasNext()){
			Sheet sheet = sheetItr.next();
			if(sheet.getSheetName() != null && eaiCode.equalsIgnoreCase(sheet.getSheetName().split("_")[0])){
				index = workbook.getSheetIndex(sheet);
			}
		}
		
		return index;
		
	}
	
	public static void generateAppInfo(Sheet sheet, PDFFileDetails pdfFileDetails) throws Exception{
		
		int colIndex = 0;
		int rowIndex =0;
		List<String> header = new ArrayList<String>();
		//Map<String,String> serverDetailsMap = new HashMap<String,String>();
		DataFormatter dataFormatter = new DataFormatter();
		
		//Get the row and cell count for iteration
		int rowsCount = sheet.getPhysicalNumberOfRows();
		int cellsCount = sheet.getRow(1).getPhysicalNumberOfCells();;
		
		Iterator<Row> rowItr = sheet.rowIterator();
		
		String sheetName  = sheet.getSheetName();
		//String eaiCode = sheetName.split("_")[0];
		
		//serverDetailsMap.put("AppName", sheetName);
		
		//Get the headers in the list
		
		//Row row1 = sheet.getRow(0);
		
		/*for(int i=0;i<row1.getPhysicalNumberOfCells();i++){
			String cellValue = dataFormatter.formatCellValue(row1.getCell(i));
			header.add(cellValue);
			
			System.out.println(header.toString());
		}*/
		
		for(int i=1;i<rowsCount;i++){
			ExcelSheetDetails sheetDetails = new ExcelSheetDetails();
			Row row = sheet.getRow(i);
			System.out.println("Row Count"+ row.getPhysicalNumberOfCells());
			for(int j=0; j<row.getPhysicalNumberOfCells();j++ ){
				
				String cellValue = dataFormatter.formatCellValue(row.getCell(j));
				//serverDetailsMap.put(header.get(j), cellValue);
				switch(j){
					case 0: sheetDetails.setEnvironment(cellValue);
					break;
					case 1: sheetDetails.setWebServerName(cellValue);
					break;
					case 2: sheetDetails.setAppServerName(cellValue);
					break;
					case 3: sheetDetails.setClusterName(cellValue);
					break;
				}
				//sheetDetails.setAppServerName(cellValue);
				System.out.println("cellValue " + cellValue);
				//System.out.println(serverDetailsMap.toString());
				
			}
			
			GenerateWordDocument.generateWord(sheetDetails,pdfFileDetails);
			
		}
		
	
		
		
	}

}
