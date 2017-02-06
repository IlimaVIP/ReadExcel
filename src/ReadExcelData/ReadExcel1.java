package ReadExcelData;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel1 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		File file = new File("C:\\Users\\matikhan89\\Desktop\\TestDATA\\TestData.xlsx");
		FileInputStream fis= new FileInputStream(file);
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		wb.getSheetAt(0);
		XSSFSheet sheet1=wb.getSheetAt(0);
		String data0 = sheet1.getRow(0).getCell(0).getStringCellValue();
		System.out.println("Data from Excel is " +data0);
		
		String data1 = sheet1.getRow(0).getCell(1).getStringCellValue();
		System.out.println("Data from Excel is " +data1);
		

		 wb.close();

        

	}

}
