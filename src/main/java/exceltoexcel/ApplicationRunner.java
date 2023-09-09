package exceltoexcel;

import exceltoexcel.serviceClasses.Extractor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;

public class ApplicationRunner {
    public static void main(String[] args) throws IOException {
        Extractor extractor = new Extractor();
        XSSFSheet sheet = extractor.prepareSheet();
        HashMap<String,String> typesOfAnalyses = extractor.prepareTypesOfAnalyses();
        HashMap<String, String> numbersOfAnalyses = extractor.prepareNumbersOfCellWithAnalyse();
        ArrayList<Sample> extractedSamples = extractor.extract(numbersOfAnalyses, typesOfAnalyses,sheet);
        for ( Sample sample:extractedSamples){
          //  System.out.println(sample.getSampleParts().toString());
         //   System.out.println(extractedSamples.size());
        }
    }

}
