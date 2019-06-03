package TestWithExcelFiles;

import org.apache.tika.parser.txt.CharsetDetector;

import java.io.*;

public class Charset {
    public static void main(String[] args) throws IOException {
        String filename = "/Users/MacbookPro/Desktop/Bash_commands/Files/CNCRS.xsd";
        BufferedReader reader = new BufferedReader(new FileReader(filename));
        CharsetDetector detector = new CharsetDetector();
        try {
            detector.setText(reader.readLine().getBytes());
        } catch (IOException e) {
            e.printStackTrace();
        }
        detector.detect();
        System.out.println(detector.detect());
    }
}
