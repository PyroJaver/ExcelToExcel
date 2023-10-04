package exceltoexcel;

import exceltoexcel.serviceClasses.Extractor;
import exceltoexcel.serviceClasses.Writer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

public class ApplicationRunner {
    public static void main(String[] args) throws IOException {
        Extractor extractor = new Extractor();
        Writer writer = new Writer();
        Utils utils = new Utils();
        XSSFSheet sheet = extractor.prepareSheetToRead();
       // XSSFWorkbook workbook = writer.prepareWorkbookToWrite();

        HashMap<String, String> typesOfAnalyses = utils.prepareProperties("analyses_reading_names.properties");
        HashMap<String, String> numbersOfCellsToRead = utils.prepareProperties("analyses_reading_numbers_of_cells.properties");
        HashMap<String, String> numbersOfCellsToWrite = utils.prepareProperties("analyses_writing_numbers_of_cells.properties");
        ArrayList<Sample> extractedSamples = extractor.extract(numbersOfCellsToRead, typesOfAnalyses,sheet);
    //    System.out.println(extractedSamples.get(1).getSampleParts().toString());
        writer.writeToExcel(numbersOfCellsToWrite, extractedSamples);
      //  System.out.println(extractedSamples.get(3).getSampleParts().toString()+"khui");
     //   for (Sample sample:extractedSamples){
          //  System.out.println(sample.getSampleParts().toString());
         //   System.out.println(extractedSamples.size());
    //    }
    }

}
