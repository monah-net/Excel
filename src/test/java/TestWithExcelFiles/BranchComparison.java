package TestWithExcelFiles;

import org.apache.poi.hslf.model.Sheet;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

public class BranchComparison {
    public static void main(String[] args) {
        String excelFile = "/Users/olegsolodovnikov/Desktop/BranchComparison/FCRS_DATA.xls";
        try {
            FileInputStream fileInput = new FileInputStream(excelFile);
            HSSFWorkbook wb = new HSSFWorkbook(fileInput);
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                HSSFSheet sheet = wb.getSheetAt(i);
                System.out.println(sheet.getSheetName());
                for (int j = sheet.getFirstRowNum(); j < sheet.getLastRowNum(); j++) {
                    HSSFRow row = sheet.getRow(j);
                    for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {
                        HSSFCell cell = row.getCell(k);
                        if (cell.getCellType() == cell.CELL_TYPE_BLANK){
                            System.out.println("null");
                        }else if(cell.getCellType() == cell.CELL_TYPE_BOOLEAN){
                            System.out.println(cell.getBooleanCellValue());
                        } else if (cell.getCellType() == cell.CELL_TYPE_ERROR){
                            System.out.println(cell.getErrorCellValue());
                        } else if(cell.getCellType() == cell.CELL_TYPE_NUMERIC){
                            System.out.println(cell.getNumericCellValue());
                        } else if (cell.getCellType() == cell.CELL_TYPE_FORMULA){
                            System.out.println(cell.getCellFormula());
                        } else if (cell.getCellType() == cell.CELL_TYPE_STRING){
                            System.out.println(cell.getStringCellValue());
                        }
                    }
                }
            }
        }
        catch (Exception e){
            System.out.println(e.getMessage());
        }
    }
}
