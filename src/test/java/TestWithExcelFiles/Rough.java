package TestWithExcelFiles;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;

public class Rough {
        public static void main(String args[]) throws IOException{
            String AddressOfxlsxFile = "/Users/MacbookPro/Desktop/ApacheExcelFiles/050-FTC-002-A.xlsx";
            FileInputStream fin = new FileInputStream(AddressOfxlsxFile);
            XSSFWorkbook wb = new XSSFWorkbook(fin);
            XSSFSheet sheet = wb.getSheetAt(1);
//            System.out.println(wb.getSheetAt(1));
            int RowCountWithNullValue=0, RowCountWithoutNullValue=0;
//            System.out.println(sheet.getRow(8));
            for (int i=0;i<1000;i++){
                if (sheet.getRow(i)==null)
                    RowCountWithNullValue++;
                else{
                    RowCountWithoutNullValue++;
                    System.out.println(sheet.getRow(8).getCell(1));
                }
            }
            System.out.println("The last row number" + sheet.getLastRowNum());
            System.out.println("Count rows with null values " + RowCountWithNullValue+", Count rows without null values "+RowCountWithoutNullValue);
        }
    }