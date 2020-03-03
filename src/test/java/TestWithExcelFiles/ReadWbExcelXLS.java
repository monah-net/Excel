package TestWithExcelFiles;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ReadWbExcelXLS {
    public static void main(String[] args) throws IOException {
        ArrayList<String> allValues = new ArrayList<>();
        String fileNameTXT = "/Users/olegsolodovnikov/Desktop/fcrs_cic/templates/File.txt";
        allValues.addAll(fileToList(fileNameTXT));
        HashMap <String,String> replaceValues = new HashMap<>();
        for (int i = 0; i < allValues.size(); i++) {
            replaceValues.put(allValues.get(i).replaceAll("\t.*",""),allValues.get(i).replaceAll(".*\t",""));
        }
        List <String> replaceableValues = new ArrayList<>();
        for (int i = 0; i < allValues.size(); i++) {
            replaceableValues.add(allValues.get(i).replaceAll("\t.*",""));
        }
        String fileNamepathXLS = "/Users/olegsolodovnikov/Desktop/fcrs_cic/templates/CRS_CIC_TEST1.xlsx";
        FileInputStream finXLS = new FileInputStream(fileNamepathXLS);
        XSSFWorkbook wbXLS = new XSSFWorkbook(finXLS);
        for (int i = 0; i < wbXLS.getNumberOfSheets(); i++) {
            if(wbXLS.getSheetName(i).equals("Account Information")) {
                XSSFSheet sheetXLS = wbXLS.getSheetAt(i);
                for (int j = sheetXLS.getLastRowNum(); j < sheetXLS.getLastRowNum(); j++) {
                    XSSFRow rowXLS = sheetXLS.getRow(j);
                    for (int k = rowXLS.getFirstCellNum(); k < rowXLS.getLastCellNum(); k++) {
                        XSSFCell cellXLS = rowXLS.getCell(k);
                        if (cellXLS.CELL_TYPE_STRING == 1) {
                            if (valueExists(cellXLS.getStringCellValue(), replaceableValues)) {
                                cellXLS.setCellValue(replaceValues.get(cellXLS.getStringCellValue()));
                            }
                        }
                    }
                }
            }
            else{
                continue;
            }
        }
        FileOutputStream xlsOutPutStream = new FileOutputStream("/Users/olegsolodovnikov/Desktop/fcrs_cic/templates/CRS_CIC_TEST1UPD.xlsx");
        wbXLS.write(xlsOutPutStream);
        xlsOutPutStream.close();
    }
    public static List <String> fileToList(String fileNamePath) throws IOException {
        FileReader fileReader = new FileReader(fileNamePath);
        BufferedReader reader = new BufferedReader(fileReader);
        List<String> list = new ArrayList<>();
        String text;
        while ((text = reader.readLine()) != null) {
            list.add(text);
        }
        return list;
    }
    public static boolean valueExists (String checkedValueString,List<String>checkedValuesArr){
        boolean result = false;
        for (int i = 0; i < checkedValuesArr.size(); i++) {
            if(checkedValueString.equals(checkedValuesArr.get(i))){
                result = true;
                break;
            }
        }
        return result;
    }
}