package main;

import java.util.HashMap;
import java.util.List;

import org.apache.logging.log4j.LogManager; 
import org.apache.logging.log4j.Logger;

public class Parser {

	private static final Logger logger = LogManager.getLogger(Parser.class);
	
	private static String templatePath;
	private static String outputPath;
	private static String pdfFilePath;
	private static String sheetName;
	private static String columnLetter;

	public static void main(String[] args) {
		
		logger.info("---START PARSER---");

//		templatePath = "C:\\Users\\T480S003\\Desktop\\textparser\\textparser.xlsx";
//		outputPath = "C:\\Users\\T480S003\\Desktop\\textparser\\output";
//		pdfFilePath = "C:\\Users\\T480S003\\Desktop\\textparser\\RHEL.pdf";
//		sheetName = "8.7";
//		columnLetter = "I";
		
		templatePath = args[0];
		outputPath = args[1];
		pdfFilePath = args[2];
		sheetName = args[3];
		columnLetter = args[4];
		
		Helper helper = new Helper();
		
		List<String> wordsToFind = helper.getWordsToFind(templatePath, sheetName, columnLetter); // look words from excel
		
		HashMap<String, Integer> wordOccurenceList = helper.getWordsOccurence(pdfFilePath, wordsToFind); // look up word occurence in PDF
		
		helper.printToTemplate(templatePath, wordOccurenceList, outputPath); // print results 
		
		logger.info("---END PARSER---");
	}

}
