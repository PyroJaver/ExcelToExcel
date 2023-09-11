package exceltoexcel;

import exceltoexcel.serviceClasses.Extractor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

public class ApplicationRunner {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        Extractor extractor = new Extractor();
        XSSFSheet sheet = extractor.prepareSheetToRead();
        HashMap<String,String> typesOfAnalyses = extractor.prepareTypesOfAnalyses();
        HashMap<String, String> numbersOfAnalyses = extractor.prepareNumbersOfCellWithAnalyse();
        ArrayList<Sample> extractedSamples = extractor.extract(numbersOfAnalyses, typesOfAnalyses,sheet);
      //  System.out.println(extractedSamples.get(3).getSampleParts().toString()+"khui");
     //   for (Sample sample:extractedSamples){
          //  System.out.println(sample.getSampleParts().toString());
         //   System.out.println(extractedSamples.size());
    //    }
    }

}
