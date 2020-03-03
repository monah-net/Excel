package TestWithExcelFiles;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;

public class ReadAllFiledFromFolderXLS {
    public static void main(String[] args) throws IOException {

        File folder = new File("/Users/olegsolodovnikov/Desktop/BranchComparison/");
        File[] listOfFiles = folder.listFiles();
        FileWriter result = new FileWriter ("/Users/olegsolodovnikov/Desktop/BranchComparison/AllProjects.txt");
        BufferedWriter writer = new BufferedWriter(result);
        for (File file : listOfFiles){
            //System.out.println(file.toString().replaceAll("^[a-zA-Z0-9\\/\\_]+\\.xls$","xls"));
            if (file.isFile() && file.toString().replaceAll("^[a-zA-Z0-9\\/\\_]+\\.xls$","xls").equals("xls")) {
                FileInputStream fin = new FileInputStream(folder.toString() + "/" + file.getName().toString());
                HSSFWorkbook wb = new HSSFWorkbook(fin);
                result.write(wb.getSheetAt(0).getRow(5).getCell(2) + "\t" + wb.getSheetAt(0).getRow(5).getCell(3) + "\t" + wb.getSheetAt(0).getRow(6).getCell(2) + "\t" + wb.getSheetAt(0).getRow(6).getCell(3) + "\n");
            }
        }
        writer.close();
    }
}
