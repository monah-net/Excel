package TestWithExcelFiles;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class MyTests {
    public static void main(String[] args) throws IOException {
        String AddressOfxlsxFile = "/Users/MacbookPro/Desktop/ApacheExcelFiles/050-FTC-002-A.xlsx";
        FileInputStream fin = new FileInputStream(AddressOfxlsxFile);
        XSSFWorkbook wb = new XSSFWorkbook(fin);
        XSSFSheet sheet = wb.getSheetAt(0);
        String sheetName = "CRS_Persons";
        String sheetNameReplaced = sheetName.replace('_',' ');
        System.out.println("Before: " + sheetNameReplaced + " After : " + sheetName);
        double doub = 555000.0;
        String string = String.valueOf(doub);
        System.out.println("New string " + string);
    }
}
