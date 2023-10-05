package exceltoexcel;

import exceltoexcel.serviceClasses.Extractor;
import exceltoexcel.serviceClasses.Writer;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

public class ApplicationRunner {
    public static void main(String[] args) throws IOException {
        Extractor extractor = new Extractor();
        Writer writer = new Writer();
        Utils utils = new Utils();
        XSSFSheet sheet = extractor.prepareSheetToRead();
        HashMap<String, String> typesOfAnalyses = utils.prepareProperties("analyses_reading_names.properties");
        HashMap<String, String> numbersOfCellsToRead = utils.prepareProperties("analyses_reading_numbers_of_cells.properties");
        HashMap<String, String> numbersOfCellsToWrite = utils.prepareProperties("analyses_writing_numbers_of_cells.properties");
        ArrayList<Sample> extractedSamples = extractor.extract(numbersOfCellsToRead, typesOfAnalyses,sheet);
        writer.writeToExcel(numbersOfCellsToWrite, extractedSamples);

    }

}

