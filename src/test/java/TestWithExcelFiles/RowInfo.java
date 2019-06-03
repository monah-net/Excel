package TestWithExcelFiles;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;

public class RowInfo {
    public static void main(String args[]) throws IOException{
        String AddressOfxlsxFile = "/Users/MacbookPro/Desktop/ApacheExcelFiles/AEOIFATCA_1291/050-FTC-002-A_GuoTian.xlsx";
        FileInputStream fin = new FileInputStream(AddressOfxlsxFile);
        XSSFWorkbook wb = new XSSFWorkbook(fin);
        XSSFSheet sheet = wb.getSheetAt(0);
//        System.out.println(sheet.getRow(12));
            /*System.out.println(sheet.getRow(8).getCell(1).getCellType());
            System.out.println(sheet.getRow(8));*/
        int NullRows = 0;
        int CellsNull = 0;
        for (int i = 8; i <= sheet.getLastRowNum() ; i++) {
            if(sheet.getRow(i) == null){
                NullRows++;
            }
        }
        for (int j = 0; j <= sheet.getLastRowNum() ; j++) {
            for (int k = 0; k <= 9 ; k++) {
                if(sheet.getRow(j).getCell(k) == null){
//                    System.out.println("Row :" + sheet.getRow(j));
                    CellsNull++;
                }
            }
        }
//        System.out.println("Row 8" + sheet.getRow(20));
        System.out.println("The last row number" + sheet.getLastRowNum());
        System.out.println("NullRows" + NullRows);
        System.out.println("NullCells" + CellsNull);
    }
}