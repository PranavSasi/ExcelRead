package excelReadSample;

import java.io.IOException;

public class ExcelMain {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String x= ExcelSub.readStringData(0, 0);
		double y= ExcelSub.readIntegerData(0, 1);
		System.out.println(x);
		System.out.println(y);
	}

}
