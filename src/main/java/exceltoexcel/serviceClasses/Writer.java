package exceltoexcel.serviceClasses;

import exceltoexcel.Sample;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Writer {
    int currentRow = 7;
    public XSSFWorkbook prepareSheetToWrite() throws IOException {
        //получаем доступ к листу эксель
        File fileToWrite = new File("C:\\Users\\kekec\\Desktop\\HallOffice2023Test.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();
  //      XSSFSheet sheet = workbook.getSheetAt(0);
        return workbook;
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
    public void writeToExcel(HashMap<String, String> numbersOfCellsToWrite, Sample sampleToWrite, XSSFWorkbook workbook){

        HashMap<String, String> sampleParts = sampleToWrite.getSampleParts();
        Set<Map.Entry<String, String>> samplePartsSet = sampleParts.entrySet();
        XSSFSheet sheetToWrite = workbook.getSheetAt(0);
        Row row = sheetToWrite.createRow(currentRow);
;

        for(Map.Entry<String, String> samplePart:samplePartsSet){
                Cell currentCell = row.createCell(Integer.parseInt(numbersOfCellsToWrite.get(samplePart.getKey())),
                        CellType.STRING);
                currentCell.setCellValue(samplePart.getValue());
         //   }
        }
        currentRow++;
    }
}
