package TestWithExcelFiles;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class FileReading {
    public static void main(String[] args) throws IOException {
        Map<String, String> mapa1 = new HashMap<String, String>();
        Map<String, String> mapa2 = new HashMap<String, String>();
        String AddrOfxlsxFile = "/Users/olegsolodovnikov/Desktop/xlsFiles/AEOI_comparison.xls";
        FileInputStream fin = new FileInputStream(AddrOfxlsxFile);
        HSSFWorkbook wb = new HSSFWorkbook(fin);
        HSSFSheet sheet = wb.getSheet("Portfolio");
        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
            HSSFRow row = sheet.getRow(i);
            if (row != null) {
                for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
                    HSSFCell cellNode = row.getCell(1);
                    if (cellNode != null) {
                        HSSFCell cellCond1 = row.getCell(3);
                        HSSFCell cellCond2 = row.getCell(4);
                        Pattern pattern = Pattern.compile("<CONDITION>.*");
                        if (cellCond1 != null && cellCond2 != null) {
                            Matcher m = pattern.matcher(cellCond1.getStringCellValue());
                            Matcher m2 = pattern.matcher(cellCond2.getStringCellValue());
                            while (m.find() && m2.find()) {
                                mapa1.put(cellNode.getStringCellValue(), cellCond1.getStringCellValue());
                                mapa2.put(cellNode.getStringCellValue(), cellCond2.getStringCellValue());
                            }
                        }
                    }
                }
            }
            for (Map.Entry<String, String> pair : mapa1.entrySet()
            ) {
                String key = pair.getKey();
                String value = pair.getValue();
                System.out.println(key + ":" + value);
            }
        }
    }
}
