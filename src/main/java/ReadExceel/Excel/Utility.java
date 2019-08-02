package ReadExceel.Excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Utility {

	public static void main(String[] args) throws IOException {
		String filePath ="C:\\old d drive\\nwb\\08072019";
		String name="Control File 2019-07-08.xlsx";
		utilitycl(filePath,name);
		

	}
	public static void utilitycl(String filePath,String name) throws IOException
	{
		
		File file =new File(filePath +"\\"+name);
		System.out.print(file);
		FileInputStream fis =new FileInputStream(file);
		XSSFWorkbook wb= new XSSFWorkbook(fis);
		XSSFSheet sheet =wb.getSheet("Control File");
		int rowcount = sheet.getLastRowNum()-sheet.getFirstRowNum();
		for(int i=0;i<=rowcount;i++)
		{
			XSSFRow row=sheet.getRow(i);
			System.out.println();
			for(int j=0;j<row.getLastCellNum();j++)
			{
				
				System.out.print(row.getCell(j).getStringCellValue());
				
			}
			System.out.println();
		}
	}

}
