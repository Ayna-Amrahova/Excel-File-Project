package com.project.ExcelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkingExcelFile {
	
    public static void main( String[] args ) throws Exception {
    	
    	File file = new File("excelFile.xlsx");
    	
   
    	//checking file 
    	
    	if (file.isFile() && file.exists()) {
			System.out.println("excelFile is open");
		} else {
			System.out.println("file not found or can't open");
		}
    	
    	System.out.println();

    	
    	//writing data to excel file
    	
        XSSFWorkbook workbook = new XSSFWorkbook(); 
        XSSFSheet spreadsheet = workbook.createSheet("Hero data"); 
        XSSFRow rowWrite; 
 
        Map<String, Object[]> heroData = new TreeMap<String, Object[]>(); 
        heroData.put("1", new Object[] {"Name", "Age"}); 
        heroData.put("2", new Object[] { "Loki", "1054"}); 
        heroData.put("3", new Object[] { "Deadpool", "28"}); 
  
        Set<String> keyid = heroData.keySet(); 
        int rowid = 0; 
        
        for (String key : keyid) { 
        	rowWrite = spreadsheet.createRow(rowid++); 
            Object[] objectArr = heroData.get(key); 
            int cellid = 0; 
            for (Object obj : objectArr) { 
                Cell cellWrite = rowWrite.createCell(cellid++); 
                cellWrite.setCellValue((String)obj); 
            } 
        } 
        
        FileOutputStream out = new FileOutputStream(file);	
        workbook.write(out); 
        out.close(); 
    	
    	

    	//Reading data from excel file
    	
     	FileInputStream fip = new FileInputStream(file);	
    	XSSFSheet sheet = workbook.getSheetAt(0);  
    	FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();  
    	for(Row rowRead: sheet){  
    		for(Cell cellRead: rowRead){  
    			switch(formulaEvaluator.evaluateInCell(cellRead).getCellType()){  
    				case Cell.CELL_TYPE_NUMERIC:
    					System.out.print(cellRead.getNumericCellValue()+ "\t\t");   
    					break;  
    				case Cell.CELL_TYPE_STRING:   
    					System.out.print(cellRead.getStringCellValue()+ "\t\t");  
    					break;  
    				}  
    		}  
    	System.out.println();  
    	} 
    	workbook.close();
    }
}
