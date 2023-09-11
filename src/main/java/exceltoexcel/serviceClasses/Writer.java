package exceltoexcel.serviceClasses;

import exceltoexcel.Sample;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class Writer {
    int currentRow = 7;
    public XSSFSheet prepareSheetToWrite() throws IOException {
        //получаем доступ к листу эксель
        FileInputStream file = new FileInputStream(new File("C:\\Users\\kekec\\Desktop\\HallOffice2023Test.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        return sheet;
    }
    public HashMap<String, String> prepareNumbersOfColumnsWrite() throws IOException {

        //получаем доступ к файлу пропертис с именами анализов
        String rootPath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
        String appConfigPath = rootPath + "analyses_writing_numbers_of_cells.properties";
        Properties readingProperties = new Properties();
        readingProperties.load(new FileInputStream(appConfigPath));
        HashMap<String, String> numbersOfColumnsToWrite = new HashMap<>();
        //считываем все записи из пропертис и заполняем хэшмапу с типами испытаний
        for (Map.Entry<Object, Object> propertiesEntrySet : readingProperties.entrySet()) {
            numbersOfColumnsToWrite.put((String) propertiesEntrySet.getKey(),
                    (String) propertiesEntrySet.getValue());
        }
        // System.out.println(typesOfAnalyses.toString());
        return numbersOfColumnsToWrite;
    }
    public void writeToExcel(HashMap<String, String> numbersOfCellsToWrite, Sample sampleToWrite, XSSFSheet sheetToWrite){

        HashMap<String, String> sampleParts = sampleToWrite.getSampleParts();
        Set<Map.Entry<String, String>> samplePartsSet = sampleParts.entrySet();
        Row row = sheetToWrite.createRow(currentRow);
        System.out.println(numbersOfCellsToWrite.toString());

        for(Map.Entry<String, String> samplePart:samplePartsSet){
            System.out.println(sampleParts.toString());
          //  if(!Objects.equals(numbersOfCellsToWrite.get(samplePart.getKey()), "")) {
                Cell currentCell = row.createCell(Integer.parseInt(numbersOfCellsToWrite.get(samplePart.getKey())));
                currentCell.setCellValue(samplePart.getValue());
         //   }
        }
        currentRow++;
    }
}
