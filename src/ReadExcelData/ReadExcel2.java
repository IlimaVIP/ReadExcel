package ReadExcelData;

import lib.ExcelDataConfig;

public class ReadExcel2 {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		ExcelDataConfig excel=new ExcelDataConfig("C:\\Users\\matikhan89\\Desktop\\TestDATA\\TestData.xlsx");
		System.out.println(excel.getData(1, 0, 1));

	}

}
