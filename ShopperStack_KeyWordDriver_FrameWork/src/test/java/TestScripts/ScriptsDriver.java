package TestScripts;

import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;

import Excel_Reader.ExcelDataReader;
import Utility_Library.KeywordsUtility;

public class ScriptsDriver {
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		KeywordsUtility keywords_Utility = new KeywordsUtility();
		
		ArrayList<String> alltestCaseNames = ExcelDataReader.getTestCaseNames();
			
		for(String testCasename : alltestCaseNames)
		{
			
			ArrayList<String> allKeyWords = ExcelDataReader.getKeyWords(testCasename);
			
			int cellNumber=1;
			for(String keyword : allKeyWords)
			{
				switch (keyword) {
				case "launchBrowser":
					keywords_Utility.launchBrowser();
					break;
				case "enterURL":
					keywords_Utility.enterUrl(ExcelDataReader.getTestData(testCasename, cellNumber));
					break;
				case "click":
					keywords_Utility.click(ExcelDataReader.getLocators(testCasename, cellNumber));
					break;
				case "enterData":
					keywords_Utility.enterData(ExcelDataReader.getLocators(testCasename, cellNumber), ExcelDataReader.getTestData(testCasename, cellNumber));
					break;
				case "verifyByMessage":
					keywords_Utility.verify(ExcelDataReader.getLocators(testCasename, cellNumber), testCasename);
					keywords_Utility.closeBrowser();
					break;
				default:
					break;
				}
				cellNumber++;
			}
			
		}
	}
}
