package TestWithExcelFiles;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class FloatCompareTests {
    public static void main(String[] args) throws IOException {
         String xlsFile = "/Users/olegsolodovnikov/Desktop/exprTESTS/oracle2.xls";
         File file = new File(xlsFile);
         FileInputStream stream = new FileInputStream(file);
         HSSFWorkbook wb = new HSSFWorkbook(stream);
         Sheet sheet = wb.getSheetAt(0);
        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            float flt = ((float) row.getCell(4).getNumericCellValue());
            System.out.println(flt);
        }
    }
}
