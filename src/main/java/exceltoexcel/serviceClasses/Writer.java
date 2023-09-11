package exceltoexcel.serviceClasses;

import exceltoexcel.Sample;
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
    //эта переменная указывает, с какой строки начнётся запись в файл
    int currentRow = 7;
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
        return numbersOfColumnsToWrite;
    }
    public void writeToExcel(HashMap<String, String> numbersOfCellsToWrite, ArrayList<Sample> samplesToWrite) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheetToWrite = workbook.createSheet("HallOffice2023TestSheet");
        for (Sample sampleToWrite : samplesToWrite) {
          //  System.out.println(sampleToWrite.getSampleParts().toString());
            HashMap<String, String> sampleParts = sampleToWrite.getSampleParts();
            Set<Map.Entry<String, String>> samplePartsSet = sampleParts.entrySet();
            Row row = sheetToWrite.createRow(currentRow);
            for (Map.Entry<String, String> samplePart : samplePartsSet) {
                Cell currentCell = row.createCell(Integer.parseInt(numbersOfCellsToWrite.get(samplePart.getKey())),
                        CellType.STRING);
                currentCell.setCellValue(samplePart.getValue());
            }
            currentRow++;

        }
        try{
            FileOutputStream outputWorkbookStream =
                    new FileOutputStream(new File("C:\\Users\\kekec\\Desktop\\HallOffice2023Test.xlsx"));
            workbook.write(outputWorkbookStream);
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }

}
