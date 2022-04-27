package org.logon;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class ExcelCreate {

	public static void main(String[] args) throws IOException {

		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();

		driver.get("http://demo.automationtesting.in/Register.html");

		WebElement skillsDown = driver.findElement(By.id("Skills"));
		Select s = new Select(skillsDown);
		List<WebElement> options = s.getOptions();
//		int size = options.size();
//		System.out.println(size);
		File file = new File("C:\\Users\\dd\\Desktop\\NewExcel.xlsx");
//		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook();

		Sheet sheet = workbook.createSheet("DATA");
		for (int i = 0; i < options.size(); i++) {

			String value = options.get(i).getText();
			Row row = sheet.createRow(i);

			Cell cell = row.createCell(0);
			cell.setCellValue(value);

		}

		FileOutputStream outputStream = new FileOutputStream(file);
		workbook.write(outputStream);
	}

}
