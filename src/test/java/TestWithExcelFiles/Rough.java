package TestWithExcelFiles;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Rough {
        public static void main(String args[]) throws IOException{
            String AddressOfxlsxFile = "/Users/MacbookPro/Desktop/ApacheExcelFiles/050-FTC-002-A.xlsx";
            FileInputStream fin = new FileInputStream(AddressOfxlsxFile);
            XSSFWorkbook wb = new XSSFWorkbook(fin);
            XSSFSheet sheet = wb.getSheetAt(0);
            System.out.println(sheet.getRow(8).getCell(1).getCellType());
            System.out.println(sheet.getRow(8).getCell(1).getStringCellValue());
            System.out.println(sheet.getRow(8));
            int RowsCreated = 0;
            int CellsCreated = 0;
            for (int i = 8; i <= sheet.getLastRowNum() ; i++) {
                if(sheet.getRow(i) == null){
                    sheet.createRow(i);
                    RowsCreated++;
                }
            }
            for (int j = 0; j <= sheet.getLastRowNum() ; j++) {
                for (int k = 0; k <= 9 ; k++) {
                    if(sheet.getRow(j).getCell(k) == null){
                        sheet.getRow(j).createCell(k,3);
                        sheet.getRow(j).getCell(k).setAsActiveCell();
                        sheet.getRow(j).getCell(k).setCellValue("");
                        CellsCreated++;
                    }
                }
            }
            System.out.println("Rows created " + RowsCreated);
            System.out.println("Cells created " + CellsCreated);
            FileOutputStream fileOut = new FileOutputStream("/Users/MacbookPro/Desktop/ApacheExcelFiles/050-FTC-002-A-3.xlsx");
            wb.write(fileOut);
            fileOut.close();
        }
    }