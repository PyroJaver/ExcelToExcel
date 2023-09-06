package exceltoexcel;

import exceltoexcel.serviceClasses.Extractor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

public class ApplicationRunner {
    public static void main(String[] args) throws IOException {
        Extractor extractor = new Extractor();
        XSSFSheet sheet = extractor.prepareSheet();
        HashMap<String,String> typesOfAnalyses = extractor.prepareTypesOfAnalyses();
        HashMap<String,String> sampleParts = extractor.extract(typesOfAnalyses,sheet);
    }

}
