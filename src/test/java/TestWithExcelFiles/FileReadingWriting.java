package TestWithExcelFiles;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class FileReadingWriting {
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
                        Pattern pattern2 = Pattern.compile("refiling.*");
                        if (cellCond1 != null && cellCond2 != null) {
                            Matcher m = pattern.matcher(cellCond1.getStringCellValue());
                            Matcher m2 = pattern.matcher(cellCond2.getStringCellValue());
                            Matcher m3 = pattern2.matcher(cellNode.getStringCellValue());
                            while (m.find() && m2.find() && !m3.find()) {
                                mapa1.put(cellNode.getStringCellValue(), cellCond1.getStringCellValue());
                                mapa2.put(cellNode.getStringCellValue(), cellCond2.getStringCellValue());
                            }
                        }
                    }
                }
            }
            FileWriter writer = new FileWriter("/Users/olegsolodovnikov/Desktop/xlsFiles/result.txt");
            BufferedWriter writerBuff = new BufferedWriter(writer);
            for (Map.Entry<String, String> pair : mapa1.entrySet()
            ) {
                String key = pair.getKey();
                String value = pair.getValue();
                writerBuff.write(ReplaceStr.replaceString(pair.getKey())+ ReplaceStr.replaceString(pair.getValue() + "\n" + "\r"));
            }
            writerBuff.close();
            FileWriter writer2 = new FileWriter("/Users/olegsolodovnikov/Desktop/xlsFiles/result2.txt");
            BufferedWriter writerBuff2 = new BufferedWriter(writer2);
            for (Map.Entry<String, String> pair : mapa2.entrySet()
            ) {
                String key = pair.getKey();
                String value = pair.getValue();
                writerBuff2.write(ReplaceStr.replaceString(pair.getKey()) + ReplaceStr.replaceString(pair.getValue()) + "\n");
            }
            writerBuff2.close();

        }
    }
}
