package TestWithExcelFiles;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashSet;

public class ExceltoExcelWriter {
    public static void main(String[] args) {
        String ExcelFileSource = "C:\\Users\\007\\Desktop\\R5TS2T.00307.SG.230-FTC-001-A.xlsx";
        String ExcelFileDst = "C:\\Users\\007\\Desktop\\050-FTC-001-A.xlsx";
        String ExcelFileResult = "C:\\Users\\007\\Desktop\\FileResult.xlsx";
        File fileResult = new File(ExcelFileResult);
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
                int rowNumber = (int) srcRow.getCell(0).getNumericCellValue();
                int columnNumber = (int) srcRow.getCell(1).getNumericCellValue();
                String sheetName = srcRow.getCell(2).getStringCellValue();
                String cellValue = srcRow.getCell(4).getStringCellValue();
                String cellType = srcRow.getCell(7).getStringCellValue();
                Sheet dstSheet = dstWb.getSheet(sheetName);
//                System.out.println("COUNTER : " + counter+ " Sheetname: " + sheetName + " Row : " + rowNumber + " Column : " + columnNumber + " Cell value : " + cellValue + " Cell type : " + cellType);
                if (dstSheet == null) System.out.println("Destination sheet does not exist");
                Row dstRow = dstSheet.getRow(rowNumber);
                System.out.println(dstRow.getCell(0).getCellType());
                if (dstRow == null) {
                    dstRow = dstSheet.createRow(rowNumber);
                    createdRows.add(dstRow);
                }
                System.out.println(dstRow.getCell(0).getCellType());
                Cell dstCell = dstRow.getCell(columnNumber);
                if(dstCell == null){
                    if (createdRows.contains(dstRow)){
                        dstCell = dstRow.createCell(rowNumber,3);
                        System.out.println("Col: " +  columnNumber + " does not exist in workbook, created");
                    }
                    else{
                        System.out.println("Col: " + columnNumber + "Row " + rowNumber + " does not exist in workbook");
                    }
//                    try{
//                        System.out.println(dstCell.getColumnIndex() +" "+ dstCell.getRowIndex() +" "+ dstCell.getCellType() +" "+ dstCell.getStringCellValue());
//                    }
//                    catch (Exception e){
//                        System.out.println(e.getMessage());
//                    }
                    if((cellValue == null) || cellValue.length() == 0){
                        dstCell.setCellType(dstCell.CELL_TYPE_BLANK);
                        dstCell.setCellValue("");
                        continue;
                    }
                   else if(dstCell.getCellType() == dstCell.CELL_TYPE_BOOLEAN){
                        dstCell.setCellValue(cellType);
                    } else if (dstCell.getCellType() == dstCell.CELL_TYPE_ERROR){
                        dstCell.setCellValue(cellValue);
                    }else if (dstCell.getCellType() == dstCell.CELL_TYPE_FORMULA){
                        try{
                            dstCell.setCellFormula(cellValue);
                        } catch (Exception e){
                            System.out.println(e.getMessage());
                        }
                    } else if (dstCell.getCellType() == dstCell.CELL_TYPE_NUMERIC){
                      dstCell.setCellValue(cellValue);
                    } else if (dstCell.getCellType() == dstCell.CELL_TYPE_BLANK){
                        if (cellType.equals("F")){
                            dstCell.setCellType(dstCell.CELL_TYPE_NUMERIC);
                            dstCell.setCellValue(cellValue);
                        } else {
                            dstCell.setCellValue(cellValue);
                        }
                    } else if (dstCell.getCellType() == dstCell.CELL_TYPE_STRING){
                        dstCell.setCellValue(cellValue);
                    } else {
                        System.out.println("Unknown cell type = " + cellType);
                    }
                }
            }
            XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator(dstWb);
            for (int sheetCounter = 0; sheetCounter < dstWb.getNumberOfSheets();sheetCounter++){
                XSSFSheet sheetName = dstWb.getSheetAt(sheetCounter);
                if (sheetName == null) continue;
                for (int rowNum = sheetName.getFirstRowNum(); rowNum < sheetName.getLastRowNum(); rowNum++){
                    XSSFRow row = sheetName.getRow(rowNum);
                    if (row == null) continue;
                    for (int colNum = row.getFirstCellNum(); colNum < row.getLastCellNum(); colNum++){
                        XSSFCell cell = row.getCell(colNum);
                        if(cell != null && cell.getCellType() == cell.CELL_TYPE_FORMULA){
                            try{
                                evaluator.evaluateFormulaCell(cell);
                            }catch (Exception e){
                                System.out.println(e.getMessage());
                            }
                        }
                    }
                }
            }
          dstWb.write(new FileOutputStream(fileResult));
            fileInput.close();
            fileOutput.close();
        }
        catch (Exception e){
            System.out.println(e.getMessage());
        }
    }
}