package com.codejava;

import java.io.File;  
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.ss.usermodel.*;  
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

public class Prvocisla {

	public static void main(String[] args) {
		String fileName;
		
		if (args.length == 0) {
			// you have to have file named "data.xlsx" in same directory as runnable *.jar
			fileName = "data.xlsx";
			
		} else {
			// or send path to *.xlsx file as parameter
			fileName = args[0];
		}
		File file = new File(fileName);
		
		// check if file exist
		if (file.exists()) {
			writePrimeNumbers(file);
		} else {
			System.out.println("File does not exist.");
		}

	}
	
	public static boolean isValid(String input) {
		// is input number?
		if (!input.matches("-?\\d+?")) {
			return false;
		}
		
		int number = Integer.parseInt(input);
		
		// is number positive?
		if (number <= 0) {
			return false;
		}
		
		return true;
	}
	
	public static boolean isPrimeNumber(int input) {
		// is number PRIME NUMBER? 
		for (int i = 2; i * i <= input; i++) {
			if ((input % i) == 0) { 
				return false; 
			}
		}
		return true;
	}
	
	public static void writePrimeNumbers(File file) {
		try {
			OPCPackage pkg = OPCPackage.open(file);
			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			XSSFSheet sheet = wb.getSheetAt(0);
			wb.close();
			
			for (Row row : sheet) {
				for (Cell cell : row) {
					// get string of column index
					String columnIndexString = CellReference.convertNumToColString(cell.getColumnIndex());
					
					// we take only data from column "B"
					if (columnIndexString.equals("B")) { 
						
						// get value of cell
						String valueString = new DataFormatter().formatCellValue(cell);
						
						// is value valid number?
						if (isValid(valueString)) {
							
							// if yes, convert it into INT
							int valueInt = Integer.parseInt(valueString);
							
							// is our INT PRIME NUMBER?
							if (isPrimeNumber(valueInt)) {
								// print PRIME NUMBER
								System.out.println(valueInt);
							}
						}
					}
				}
			}
		} catch(Exception e) {
			e.printStackTrace();  
		}
	}

}
