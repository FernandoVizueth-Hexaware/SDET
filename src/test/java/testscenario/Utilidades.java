package testscenario;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;;

public class Utilidades {
	
	public Utilidades(){};
	
	public WebDriver driver;
	public WebElement WebEl;
	public String stepName;

	public boolean LoadSite(String strUrl, String navegador) {
		try {
			switch (navegador) {
			case "Chrome":
				System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
				driver = new ChromeDriver();
				break;

			case "FireFox":
				System.setProperty("webdriver.gecko.driver", "geckodriver.exe");
				driver = new FirefoxDriver();
				break;

			case "InternetExplorer":
				System.setProperty("webdriver.ie.driver", "IEDriverServer.exe");
				driver = new InternetExplorerDriver();
				break;

			default:
				break;
			}
			stepName = "Cargar el sitio con la dirección: " + strUrl;
			driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
			driver.manage().window().maximize();
			driver.navigate().to(strUrl);
			return true;
		} catch (Exception e) {
			System.out.println(e.toString());
			return false;
		}
	}

	public boolean ValidateWebObjectIsVisible(String propXpath) {
		if (this.BuildWebObject(propXpath)) {
			if (WebEl.findElement(By.xpath(propXpath)).isEnabled()
					&& WebEl.findElement(By.xpath(propXpath)).isDisplayed()) {
				this.WaitTime(1000);
				return true;
			} else
				return false;
		} else
			return false;
	}

// Build web object
	public boolean BuildWebObject(String propXpath) {
		WebDriverWait myWaitDriver = new WebDriverWait(driver, 3);
		try {
			myWaitDriver.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(propXpath)));
			driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
			this.WaitTime(200);
			WebEl = driver.findElement(By.xpath(propXpath));
			// Highlight element
//			JavascriptExecutor js = (JavascriptExecutor) driver;
//			js.executeScript("arguments[0].setAttribute('style', 'background: #FDFF47; border: 2px solid #000000;');", WebEl);
//			try {
//				Thread.sleep(500);
//			} catch (InterruptedException e) {
//				System.out.println(e.getMessage());
//			}
//			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white');", WebEl);

			return true;
		} catch (TimeoutException toe) {
			System.out.println(toe.toString());
			return false;
		}
	}

	public void WaitTime(int intWaitTimeMili) {
		try {
			Thread.sleep(intWaitTimeMili);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

	public boolean SetValue(WebElement propXpath, String strDataToBeSet) {
		if (propXpath.isEnabled()) {
			propXpath.clear();
			propXpath.sendKeys(strDataToBeSet);
			this.WaitTime(200);
			return true;
		} else
			return false;

	}

	public boolean ClickOnWebElement(WebElement ele) {
		if (!ele.isSelected()) {
			ele.click();
			this.WaitTime(200);
			return true;
		}
		return true;

	}

	public boolean SelectDropIndex(WebElement ele, String sel) {
		Select drpDown = new Select(ele);
		drpDown.selectByIndex(2);
		return true;
	}

	public boolean SelectDropName(WebElement ele, String sel) {
		Select drpDown = new Select(ele);
		drpDown.selectByVisibleText(sel);
		return true;
	}

	public boolean GetContainsValueObject(WebElement WebEl, String sequence, String name) {
		try {
			if (WebEl.getText().contains(sequence)) {
				//HighLight.shadeElem(driver, WebEl);
				this.WaitTime(200);
				return true;
			}
		} catch (Exception e) {
			System.out.println("Fail Case. Function contains get value " + name);
		}
		return false;
	}

// Terminate script (Fail)
	public void Terminate() {
		driver.close();
		driver.quit();
		System.exit(1);
	}

// Conclude script (Pass)
	public void ConcludeScript() {
		driver.close();
		driver.quit();
		System.exit(0);
	}

	public void TerminateScript(int flagStatus) {
		driver.close();
		driver.quit();
		if (flagStatus == 0 || flagStatus == 1) {
			System.exit(flagStatus);
		}
	}

	/*
	 * // Highlight element public static void shadeElem(WebElement element) {
	 * JavascriptExecutor js = (JavascriptExecutor) driver; js.
	 * executeScript("arguments[0].setAttribute('style', 'background: #FDFF47; border: 2px solid #000000;');"
	 * , element); try { Thread.sleep(500); } catch (InterruptedException e) {
	 * System.out.println(e.getMessage()); } js.
	 * executeScript("arguments[0].setAttribute('style','border: solid 2px white');"
	 * , element); }
	 */
	
	public static void CreateExcel() {
        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook(); 
         
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");
          
        //This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"NAME", "LASTNAME", "EMAIL", "PASSWORD", "COMPANY", "ADDRESS", "CITY", "ZIP_CODE", "MOBILE_PHONE"});
        data.put("2", new Object[] {"Fernando", "Vizueth", "vizu74@gmail.com", "Password", "Hexaware", "Dirreccion", "CDMX", "03620", "5520703031"});
          
        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("howtodoinjava_demo.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
        
        
    }
	

}
