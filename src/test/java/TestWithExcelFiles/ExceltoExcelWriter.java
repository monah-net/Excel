package TestWithExcelFiles;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.HashSet;

public class ExceltoExcelWriter {
    public static void main(String[] args) {
        String ExcelFileSource = "";
        String ExcelFileDst = "";
        try{
            FileInputStream fileInput = new FileInputStream(ExcelFileSource);
            FileInputStream fileOutput = new FileInputStream(ExcelFileDst);
            XSSFWorkbook srcWb = new XSSFWorkbook(fileInput); //Source workbook
            XSSFWorkbook dstWb = new XSSFWorkbook(fileOutput); //Destination workbook
            Sheet srcSheet = srcWb.getSheetAt(0);
            HashSet createdRows = new HashSet();
            for (int counter = srcSheet.getFirstRowNum(); counter < srcSheet.getLastRowNum(); counter++){
                Row srcRow = srcSheet.getRow(counter);
                if(srcRow == null || srcRow.getRowNum() == 0) continue;
                double rowNumber = srcRow.getCell(0).getNumericCellValue();
                int columnNumber = srcRow.getCell(1).getColumnIndex();
                String sheetName = srcRow.getCell(2).getStringCellValue();
                String cellValue = srcRow.getCell(4).getStringCellValue();
                String cellType = srcRow.getCell(7).getStringCellValue();
                
                
            }
        }
        catch (Exception e){

        }

    }
}