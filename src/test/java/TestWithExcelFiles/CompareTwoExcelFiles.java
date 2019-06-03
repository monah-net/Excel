package TestWithExcelFiles;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class CompareTwoExcelFiles {
    public static void main(String[] args) throws IOException {
        String AddrOfxlsxFile = "/Users/MacbookPro/Desktop/ApacheExcelFiles/050-FTC-002-A.xlsx";
        FileInputStream fin = new FileInputStream(AddrOfxlsxFile);
        XSSFWorkbook wb = new XSSFWorkbook(fin);

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            for (int j = 0; j < sheet.getLastRowNum(); j++) {
                XSSFRow row = sheet.getRow(j);
                for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {
                    XSSFCell cell = row.getCell(k);
                    if(cell != null && cell.getCellType() == 3){
                        System.out.println("Zagl");
                    }
                }
            }
        }
    }
}
