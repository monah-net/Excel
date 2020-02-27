package TestWithExcelFiles;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadingWbExcelXLSX {

    public static void main(String[] args) {

        try {
            File excel = new File("/Users/olegsolodovnikov/Desktop/Test_Sample_AU_AH_500K_Ind_no_data.xlsx");
            FileInputStream fis = new FileInputStream(excel);
            XSSFWorkbook book = new XSSFWorkbook(fis);
//            for (int i = 0; i < book.getNumberOfSheets(); i++) {
//                System.out.println(book.getSheetAt(i).getSheetName());
//            }
            XSSFSheet sheet = book.getSheetAt(2);
            System.out.println(sheet.getSheetName());
            System.out.println(sheet.getLastRowNum());
            for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
                XSSFRow row = sheet.getRow(i);
                for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
                    XSSFCell cell = row.getCell(j);
                    switch (cell.getCellType()) {
                        case XSSFCell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                        case XSSFCell.CELL_TYPE_BLANK:
                            System.out.println("\t");
                        case XSSFCell.CELL_TYPE_ERROR:
                            System.out.println(cell.getErrorCellValue() + "\t");
                        case XSSFCell.CELL_TYPE_FORMULA:
                            System.out.println(cell.getCellFormula() + "\t");
                        default:
                            System.out.println("Default");
                    }
                }
            }
            fis.close();
        } catch (FileNotFoundException fe) {
            fe.printStackTrace();
        } catch (IOException ie) {
            ie.printStackTrace();
        }
    }
}
