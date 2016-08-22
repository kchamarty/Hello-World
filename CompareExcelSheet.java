package com.amdocs.examples;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelEx {
	
	static int KEY_COLUMNS[] = {2,3}; 
	static int CONSIDER_SKIP_ROW_COUNT = 5;
	
static boolean doBasicValidationOfSheets(String fileName){
	
	boolean validationStatus = false;
	Workbook wb = null;
	Sheet firstSheet = null;
	Sheet secondSheet = null;
	
	try {
		wb = new XSSFWorkbook(new FileInputStream(fileName));
		firstSheet = wb.getSheetAt(0);
		secondSheet = wb.getSheetAt(1);
		
		int firstSheetRowCount = firstSheet.getPhysicalNumberOfRows();
	    int secondSheetRowCount = secondSheet.getPhysicalNumberOfRows();
	    if (firstSheetRowCount <= 1 || secondSheetRowCount <= 1){
	    	System.out.println("Row data is missing from sheet! ");
	    	return false;
	    }
	    
	    
	    int firstSheetColumnCount = firstSheet.getRow(0).getPhysicalNumberOfCells();
	    int secondSheetColumnCount = secondSheet.getRow(0).getPhysicalNumberOfCells();
	    
	    if (firstSheetColumnCount != secondSheetColumnCount ){
	    	System.out.println("Columns are not matching! Sheet1 Column Count = "+firstSheetColumnCount + ", Sheet2 Column Count = "+secondSheetColumnCount);
	    	return false;
	    }
	    
	} catch (FileNotFoundException e) {
		System.out.println("File ("+fileName+") is not present!");
		e.printStackTrace();
		return validationStatus;
	}catch (IOException e){
		System.out.println("Got error while reading file");
		return validationStatus;
	}catch(IllegalArgumentException iE){
		System.out.println("Sheet does not exist!");
		return validationStatus;
	}finally{
		try{
			wb.close();
		}catch(Exception e){}
	}

		return true;
	}

static boolean compareKeyColumns(Row a, Row b, int columnCount, int keyColumns[]) {
	
	for (int i=0; i<keyColumns.length; i++){
		
		if(!a.getCell(keyColumns[i]-1).toString().trim().equals(b.getCell(keyColumns[i]-1).toString().trim())) return false;
	}
	
	return true;
}

	public static void main(String[] args) throws Exception{
		// TODO Auto-generated method stub
		
		String fileName = "C:\\Users\\bk\\Desktop\\Tidal Blast.xlsx";
		if (!doBasicValidationOfSheets(fileName) ) {
			System.out.println("Basic Validations failed, exiting from program!");
			System.exit(1);
		}
// Do Basic Validation

		
		Workbook wb = null;
		Sheet firstSheet = null;
		Sheet secondSheet = null;
		
		
			wb = new XSSFWorkbook(new FileInputStream(fileName));
			firstSheet = wb.getSheetAt(0);
			secondSheet = wb.getSheetAt(1);
		
           int firstSheetRowCount = firstSheet.getPhysicalNumberOfRows();
           int secondSheetRowCount = secondSheet.getPhysicalNumberOfRows();
           int iterationCount;
           int firstSheetColumnCount = firstSheet.getRow(0).getPhysicalNumberOfCells();
           int secondSheetColumnCount = secondSheet.getRow(0).getPhysicalNumberOfCells();
           
           if(firstSheetRowCount >= secondSheetRowCount)	iterationCount = firstSheetRowCount;
           else iterationCount = secondSheetRowCount;
           
           System.out.println("now of rows in sheet1  = " + firstSheetRowCount);
           System.out.println("now of rows in sheet2  = " + secondSheetRowCount);
           System.out.println("now of rows for iteration  = " + iterationCount);
           System.out.println("now of cloumns in sheet1 = " + firstSheetColumnCount);
           System.out.println("now of cloumns in sheet2 = " + secondSheetColumnCount);
          
           
           // start Data Validation
           // start row iteration
           
        // Compare Cells
   		// compare key cells
   			// if key cells are not matching, check for other row till it matches / till defined count and insert one row
   			// else continue comparison with next row
   		// compare other optional cells
   			// if not matching highlight
   			// else continue
           
            for (int i=1; i<iterationCount; i++){
            	int misMatchCount = 0;
            		while (misMatchCount <= CONSIDER_SKIP_ROW_COUNT){
            			if (!compareKeyColumns(firstSheet.getRow(i),secondSheet.getRow(i+misMatchCount), firstSheetColumnCount, KEY_COLUMNS)) {
            				misMatchCount += 1;
            			}else{
            				misMatchCount = 0;
            				// Add missing Row to Sheet 3 and insert missing row to current sheet
            				
            			}
            		}
            	
            	
            	           	            	
            	// start cell iteration            	
            	
            }
            
            
            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\bk\\Desktop\\Tidal Blast.xlsx");
            wb.write(fileOut);
            fileOut.close();

	}
}
