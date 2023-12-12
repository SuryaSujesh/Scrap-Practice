package scrape;


import java.io.IOException;
import java.time.Duration;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;


import utility.ExcelUtility;

public class Wiki {

		public static void main(String[] args) throws IOException  {
		
		WebDriver driver= new ChromeDriver();
		
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.get("https://en.wikipedia.org/wiki/List_of_countries_and_dependencies_by_population");
		
		String path = "src/test/java/dataFiles/WikiData.xlsx";
		ExcelUtility excelutility = new ExcelUtility(path);
		
		excelutility.setcelldata("Sheet1", 0, 0, "Country");
		excelutility.setcelldata("Sheet1", 0, 1, "Population");
		excelutility.setcelldata("Sheet1", 0, 2, "% ge");
		excelutility.setcelldata("Sheet1", 0, 3, "Date");
		excelutility.setcelldata("Sheet1", 0, 4, "Source");
	
	     WebElement table = driver.findElement(By.xpath("//table[@class='wikitable sortable jquery-tablesorter']"));
	     int rows= table.findElements(By.xpath("tbody/tr")).size();
	    
	     for (int i=1;i<rows;i++)
	     {
	    	 String country = table.findElement(By.xpath("//tbody/tr["+i+"]/td[1]")).getText();
	    	 String population = table.findElement(By.xpath("//tbody/tr["+i+"]/td[2]")).getText();
	    	 String percentage = table.findElement(By.xpath("//tbody/tr["+i+"]/td[3]")).getText();
	    	 String date = table.findElement(By.xpath("//tbody/tr["+i+"]/td[4]")).getText();
	    	 String source = table.findElement(By.xpath("//tbody/tr["+i+"]/td[5]")).getText();
		    
	    	 excelutility.setcelldata("Sheet1", i, 0, country);
	    	 excelutility.setcelldata("Sheet1", i, 1, population);
	    	 excelutility.setcelldata("Sheet1", i, 2, percentage);
	    	 excelutility.setcelldata("Sheet1", i, 3, date);
	    	 excelutility.setcelldata("Sheet1", i, 4, source);
	    	 
	    	
	     }  
	     
	
}}
