package exceltoexcel.serviceClasses;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class Extractor {


    public XSSFSheet prepareSheet() throws IOException {
        //получаем доступ к листу эксель
        FileInputStream file = new FileInputStream(new File("C:\\Users\\kekec\\Desktop\\HallOfficeA2023.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        return sheet;


        }
    public HashMap<String,String> prepareTypesOfAnalyses() throws IOException {
        //получаем доступ к файлу пропертис
        String rootPath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
        String appConfigPath = rootPath + "analyses_reading_names.properties";
        Properties readingProperties = new Properties();
        readingProperties.load(new FileInputStream(appConfigPath));

        HashMap<String,String> typesOfAnalyses = new HashMap<>();

        //считываем все записи из пропертис и заполняем хэшмапу с типами испытаний
        for (Map.Entry<Object, Object> propertiesEntrySet: readingProperties.entrySet()){
            typesOfAnalyses.put((String) propertiesEntrySet.getKey(), (String) propertiesEntrySet.getValue());
        }
        System.out.println(typesOfAnalyses.toString());
        return typesOfAnalyses;
    }
    public HashMap<String,String> extract(HashMap<String,String> typesOfAnalysis, XSSFSheet sheet){
        HashMap<String,String> sampleParts = new HashMap<>();
        //получаем итератор
       Iterator<Row> rowIterator = sheet.rowIterator();
       rowIterator.next();
       //начинаем прохождение по листу
        while(rowIterator.hasNext()){
            Row currentRow = rowIterator.next();
            System.out.println(currentRow.getRowNum());
            Iterator<Cell> cellIterator = currentRow.cellIterator();
            Cell cell = cellIterator.next();
            System.out.println(cell.getColumnIndex());
            String cellContain = cell.getStringCellValue();
            System.out.println(cellContain);
            if (!typesOfAnalysis.containsKey(cellContain)){
                continue;
            }
            if(typesOfAnalysis.get(cellContain).equals("Sample")){
                performSampleAndTankExtraction(currentRow, sampleParts);
                continue;
            }
            if (typesOfAnalysis.get(cellContain).equals("Comments:")){
                rowIterator = performCommentExtraction(rowIterator, sampleParts);
                continue;
            }

        }


       return sampleParts;
    }

    public void performSampleAndTankExtraction(Row row, HashMap<String, String> sampleParts){
        sampleParts.put("Sample",row.getCell(2).getStringCellValue());
        sampleParts.put("Tank",row.getCell(4).getStringCellValue());
    }

    public Iterator<Row> performCommentExtraction(Iterator<Row> rowIterator, HashMap<String, String> sampleParts){
        StringBuilder comment = new StringBuilder();
        Row row = rowIterator.next();
        do{int cellCounter = 0;
            while(!rowIterator.next().getCell(cellCounter).getStringCellValue().isEmpty()){
                comment.append(row.getCell(cellCounter).getStringCellValue()).append(" ");}
        } while (!rowIterator.next().getCell(0).getStringCellValue().isEmpty());
        sampleParts.put("Comments:", comment.toString());
        System.out.println(sampleParts.get("Comments:"));
      return rowIterator;
    }
}
