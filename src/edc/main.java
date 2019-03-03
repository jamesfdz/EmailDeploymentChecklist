package edc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.logging.FileHandler;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class main {
	
	public static final String URL = "https://login.veevavault.com/";
	 public static Logger logger = Logger.getLogger("MyLog");   
	 public static FileHandler fh;
	 public static final List<String> EFNames = new ArrayList<String>();
	 public static final List<String> EFNum = new ArrayList<String>();
	 public static final List<String> REFDOCNAME = new ArrayList<String>();
	 public static final List<String> REFDOCNUM = new ArrayList<String>();
	 
	 public static String FragmentName;
	public static String FragmentNum;
	public static String ReferenceDocName;
	public static String ReferenceDocNum;

	public static void main(String[] args) throws SecurityException, IOException {
		
		try {
			System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
			WebDriver driver = new ChromeDriver();
	        driver.manage().window().maximize();
	        driver.get(URL);
	        WebElement usernameField = driver.findElement(By.id("j_username"));
	        
	      //takes input from excel
	        FileInputStream fs = new FileInputStream("input.xlsx");
	        XSSFWorkbook inputWorkbook = new XSSFWorkbook(fs);
	        Sheet inputSheet =  inputWorkbook.getSheet("Sheet1");
	        Row row1 = inputSheet.getRow(0);
	        Cell cell0 = row1.getCell(0);
	        String username = cell0.getStringCellValue();
	        System.out.println(username);
	        usernameField.sendKeys(username);
	        usernameField.sendKeys(Keys.ENTER);
	        
	        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.id("j_password")));
	        
	        WebElement passwordField = driver.findElement(By.id("j_password"));
	        
	        Row row2 = inputSheet.getRow(1);
	        Cell cell1 = row2.getCell(0);
	        String password = cell1.getStringCellValue();
	        passwordField.sendKeys(password);
	        System.out.println(password);
	        passwordField.sendKeys(Keys.ENTER);      
	        
	        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class='noItemsFound vv_no_results']")));
	        
	        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
	        
	        WebElement searchField = driver.findElement(By.id("search_main_box"));
	        Row row3 = inputSheet.getRow(2);
	        Cell cell2 = row3.getCell(0);
	        String searchString = cell2.getStringCellValue();
	        
	        driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	        
	        searchField.sendKeys(searchString);
	        driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	        searchField.sendKeys(Keys.ENTER);
	        System.out.println(searchString);
	        
	        inputWorkbook.close();
	        
	        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@title='Detail View']")));
	        WebElement detailView = driver.findElement(By.xpath("//*[@title='Detail View']"));
	        detailView.click();
	        
	        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//*[@class='docLink vv_doc_title_link'])[1]")));
	        
	        driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	        
	        WebElement presentationTitle = driver.findElement(By.xpath("(//*[@class='docLink vv_doc_title_link'])[1]"));
	        presentationTitle.click();
	        
	        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
	        
	        new WebDriverWait(driver, 2000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@attrkey='name']")));
	        WebElement PresentationNameElement = driver.findElement(By.xpath("//*[@attrkey='name']"));
	        WebElement PresentationBinderElement = driver.findElement(By.xpath("//*[@attrkey='DocumentNumber']"));
	        
	        String PresentationNameText = PresentationNameElement.getText();
	        String PresentationBinderText = PresentationBinderElement.getText();
	                
	        FileInputStream fis = new FileInputStream("output.xlsx");
	        XSSFWorkbook outputWorkbook = new XSSFWorkbook(fis);
	        Sheet outputSheet = outputWorkbook.getSheet("Checklist");
	        
	        Cell NameCell = null;
	        Cell BinderCell = null;
	                
	        NameCell = outputSheet.getRow(7).getCell(1);
	        NameCell.setCellValue(PresentationNameText);
	        BinderCell = outputSheet.getRow(8).getCell(1);
	        BinderCell.setCellValue(PresentationBinderText);
	        
	        WebElement productInformation = driver.findElement(By.xpath("//*[@key='productInformation']"));
	        productInformation.click();
	        
	        WebElement productName = driver.findElement(By.xpath("(//*[@attrkey='product'])[1]//a"));
	        String productNameText = productName.getText();
	        
	        Cell productNameCell = null;
	        
	        productNameCell = outputSheet.getRow(2).getCell(1);
	        productNameCell.setCellValue(productNameText);
	        
	        //extracting other related docs
	        
	        WebElement otherDocEm = driver.findElement(By.xpath("//*[@key='other_pm']//span[@class='count vv_section_count']"));
	        String otherDocEmString = otherDocEm.getText();
	        System.out.println(otherDocEmString);
	        
	        if(otherDocEmString.equals("(0)")){
	        	Cell ReferenceDocNameCell = null;
	            Cell ReferenceDocNumCell = null;
	            
	            ReferenceDocNameCell = outputSheet.getRow(12).getCell(1);
	            ReferenceDocNameCell.setCellValue("N/A");
	            
	            ReferenceDocNumCell = outputSheet.getRow(13).getCell(1);
	            ReferenceDocNumCell.setCellValue("N/A");
	        }else {
	        	WebElement relatedSharedElement = driver.findElement(By.xpath("//*[@key='other_pm']"));
	            relatedSharedElement.click();
	            new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class='vv_rd_c2']//a[@class='docName doc_link veevaTooltipBound']")));
	            
	            List<WebElement> refDocEms = driver.findElements(By.xpath("//*[@class='relatedDoc vv_related_item vv_rd minProps']//a[@class='docName doc_link veevaTooltipBound']"));
	            
	            for(WebElement refDocEm : refDocEms) {
	            	System.out.println(refDocEm.getText());
	            	REFDOCNAME.add(refDocEm.getText());
	            }
	            
	            List<WebElement> refDocNums = driver.findElements(By.xpath("//*[@class='relatedDoc vv_related_item vv_rd minProps']//span[@class='docNumber vv_rd_body_3']"));
	            
	            for(WebElement refDocNum : refDocNums) {
	            	System.out.println(refDocNum.getText());
	            	REFDOCNUM.add(refDocNum.getText());
	            }
	            
	            ReferenceDocName = REFDOCNAME.toString().replace("[", "")
	                    .replace("]", "");
	            
	            ReferenceDocNum = REFDOCNUM.toString().replace("[", "")
	                    .replace("]", "");
	            
	            Cell ReferenceDocNameCell = null;
	            Cell ReferenceDocNumCell = null;
	            
	            ReferenceDocNameCell = outputSheet.getRow(12).getCell(1);
	            ReferenceDocNameCell.setCellValue(ReferenceDocName);
	            
	            ReferenceDocNumCell = outputSheet.getRow(13).getCell(1);
	            ReferenceDocNumCell.setCellValue(ReferenceDocNum);
	            
	        }
	        
	        //extracting related fragment details
	        
	        WebElement relatedEmailFragmentCheck = driver.findElement(By.xpath("//*[@key='relatedProductFragments_pm']//span[@class='count vv_section_count']"));
	        String relatedEmailFragmentCheckString = relatedEmailFragmentCheck.getText();
	        System.out.println(relatedEmailFragmentCheckString);
	        
	        if(relatedEmailFragmentCheckString.equals("(0)")){
	        	Cell FragmentNameCell = null;
	            Cell FragmentNumCell = null;
	            
	            FragmentNameCell = outputSheet.getRow(9).getCell(1);
	            FragmentNameCell.setCellValue("N/A");
	            
	            FragmentNumCell = outputSheet.getRow(10).getCell(1);
	            FragmentNumCell.setCellValue("N/A");
	        }else {
	        	WebElement relatedSharedElement = driver.findElement(By.xpath("//*[@key='relatedProductFragments_pm']"));
	            relatedSharedElement.click();
	            new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class='vv_rd_c2']//a[@class='docName doc_link veevaTooltipBound']")));
	            
	            List<WebElement> evenEFNames = driver.findElements(By.xpath("//*[@class='relatedDoc vv_related_item vv_rd minProps']//a[@class='docName doc_link veevaTooltipBound']"));
	            
	            for(WebElement evenEFName : evenEFNames) {
	            	System.out.println(evenEFName.getText());
	            	EFNames.add(evenEFName.getText());
	            }
	            
	            List<WebElement> oddEFNames = driver.findElements(By.xpath("//*[@class='relatedDoc vv_related_item vv_rd minProps odd_bg']//a[@class='docName doc_link veevaTooltipBound']"));
	            
	            for(WebElement oddEFName : oddEFNames) {
	            	System.out.println(oddEFName.getText());
	            	EFNames.add(oddEFName.getText());
	            }
	            
	            List<WebElement> evenEFNums = driver.findElements(By.xpath("//*[@class='relatedDoc vv_related_item vv_rd minProps']//span[@class='docNumber vv_rd_body_3']"));
	            
	            for(WebElement evenEFNum : evenEFNums) {
	            	System.out.println(evenEFNum.getText());
	            	EFNum.add(evenEFNum.getText());
	            }
	            
	            List<WebElement> oddEFNums = driver.findElements(By.xpath("//*[@class='relatedDoc vv_related_item vv_rd minProps odd_bg']//span[@class='docNumber vv_rd_body_3']"));
	            
	            for(WebElement oddEFNum : oddEFNums) {
	            	System.out.println(oddEFNum.getText());
	            	EFNum.add(oddEFNum.getText());
	            }
	            
	            FragmentName = EFNames.toString().replace("[", "")
	                    .replace("]", "");
	            
	            FragmentNum = EFNum.toString().replace("[", "")
	                    .replace("]", "");
	            
	            Cell FragmentNameCell = null;
	            Cell FragmentNumCell = null;
	            
	            FragmentNameCell = outputSheet.getRow(9).getCell(1);
	            FragmentNameCell.setCellValue(FragmentName);
	            
	            FragmentNumCell = outputSheet.getRow(10).getCell(1);
	            FragmentNumCell.setCellValue(FragmentNum);
	            
	        }
	        
	        driver.close();
	        
	        fis.close();
	        FileOutputStream output = new FileOutputStream(new File("output.xlsx"));
	        outputWorkbook.write(output);
	        output.close();
	        
	        outputWorkbook.close();
	        
	        JOptionPane.showMessageDialog(null, "Completed Successfully");
		}catch(Exception e){
            fh = new FileHandler("LogFile.log");  
            logger.addHandler(fh);
            SimpleFormatter formatter = new SimpleFormatter();  
            fh.setFormatter(formatter);  

           // the following statement is used to log any messages
           logger.info("" + e.getLocalizedMessage());
       } 
        
	}

}
