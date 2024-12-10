package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class Helper {
	
	private static final Logger logger = LogManager.getLogger(Parser.class);
	
	private static final String TEMPLATE_OUTPUT_SHEET = "出力";
	private static final String COLUMN_KEYWORD = "キーワード";
	private static final String COLUMN_OCCURENCE = "発生";
	
	public List<String> getWordsToFind(String filepath, String sheetName, String columnLetter) {
		List<String> listWords = new ArrayList<>();
		
		try {
			
			FileInputStream fis = new FileInputStream(filepath);
			XSSFWorkbook xlswb = new XSSFWorkbook(fis);
			
			Sheet sheet = xlswb.getSheet(sheetName); // Get the sheet name
			if (sheet == null) {
				logger.error("Sheet " + sheetName + " not found!"); 
				xlswb.close();
				return null; 
			} // end if
			
			// Convert column letter to index
            int columnIndex = columnLetterToIndex(columnLetter);
            if (columnIndex == -1) {
            	logger.error("Invalid column letter " + columnLetter + "!");
                xlswb.close();
                return null;
            } // end if
			
			Iterator<Row> rowIterator = sheet.iterator();
			
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Cell cell = row.getCell(columnIndex);				
				String strValue = "";
				
				if (cell != null) {					
					if(cell.getCellType() == CellType.STRING && cell.getStringCellValue().equals("-") || cell.getCellType() == CellType.BLANK) {
						continue; // Skip this cell
					}// end if
					
                    switch (cell.getCellType()) {                    
                        case NUMERIC:
                        	strValue = String.valueOf(cell.getNumericCellValue());
                            break;                        
                        default:
                        	strValue = cell.getStringCellValue();
                            break;
                    }
                    
                    listWords.add(strValue);
                    
                } // end if
				
			} // end while
			xlswb.close();			
		}catch (IOException e) {
			logger.error("Error while getting words to find.", e);
		}
		
		return listWords;
	}
	
	public HashMap<String, Integer> getWordsOccurence(String filePath, List<String> wordsToFind) {
		HashMap<String, Integer> wordsMap = new HashMap<>();
		
		try {			 
			 PDDocument document = PDDocument.load(new File(filePath));
			 PDFTextStripper pdfStripper = new PDFTextStripper(); 
			 String text = pdfStripper.getText(document);
			 document.close();
			 
			 for (int i = 0; i < wordsToFind.size(); i++) {				 
				 int count = countOccurrences(text, wordsToFind.get(i));				 
				 wordsMap.put(wordsToFind.get(i), count);				 
			 }
			 	 
			 
		 } catch (IOException e) {
			 e.printStackTrace(); 
		 }
		
		return wordsMap;
	}
	
	public void printToTemplate(String templateLocation, HashMap<String, Integer> words, String outPath) {
		LocalDateTime now = LocalDateTime.now();
		
		try {
			
			FileInputStream fis = new FileInputStream(templateLocation);
			XSSFWorkbook xlswb = new XSSFWorkbook(fis);
			XSSFSheet outputSheet = xlswb.getSheet(TEMPLATE_OUTPUT_SHEET);
			
			// Check if the sheet exists
			if (outputSheet == null) { 
				// Create the sheet if it does not exist
				outputSheet = xlswb.createSheet(TEMPLATE_OUTPUT_SHEET);
			}
			
			// clear contents of output sheet
			for (int i = outputSheet.getLastRowNum(); i >= outputSheet.getFirstRowNum(); i--) { 
				if (outputSheet.getRow(i) != null) { 
					outputSheet.removeRow(outputSheet.getRow(i));
				} 
			}
			
			XSSFRow row = outputSheet.createRow(0); //create row for the inputs
			
			// print label for first row
			row.createCell(0).setCellValue(COLUMN_KEYWORD);
			row.createCell(1).setCellValue(COLUMN_OCCURENCE);
			
			int idx = 0;
			// print words and details
			if (words != null && words.size() > 0) { 
				
				for (Entry<String, Integer> entry : words.entrySet()) {
					
					// print to output sheet
					row = outputSheet.createRow(idx + 1);
					row.createCell(0).setCellValue(entry.getKey());
					row.createCell(1).setCellValue(entry.getValue());
					idx++;
					
				}				                        
				
			}			
			
			fis.close();
			
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
			
			// Step 1: Check if the folder exists
			File folder = new File(outPath);
			if (!folder.exists()) { 
				// Step 2: Create the folder if it does not exist
				if (folder.mkdirs()) { 
					logger.info("Folder created successfully.");
				}else {
					logger.error("Failed to create the folder.");					
				}
			}			
			
			String tempFileName = folder + "\\textparser_" + (String) dtf.format(now) + ".xlsx";
			FileOutputStream fos = new FileOutputStream(new File(tempFileName));
			
			xlswb.write(fos);
			xlswb.close();
			
		}catch(IOException e) {
			logger.error("Error while reporting results.", e);
		}		
	}
	
	private static int columnLetterToIndex(String columnLetter) {
        int columnIndex = -1;
        if (columnLetter != null && !columnLetter.isEmpty()) {
            columnIndex = 0;
            for (int i = 0; i < columnLetter.length(); i++) {
                columnIndex *= 26;
                columnIndex += columnLetter.charAt(i) - 'A' + 1;
            }
            columnIndex--; // Convert 1-based index to 0-based index
        }
        return columnIndex;
    }
	
	private static int countOccurrences(String text, String keyword) {	
	    int count = 0; 
	    int index = 0; 
	    while ((index = text.indexOf(keyword, index)) != -1) {
	        count++; index += keyword.length(); 
	    } 
	    return count;	    
	}

}
