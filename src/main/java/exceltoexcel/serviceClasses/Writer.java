package exceltoexcel.serviceClasses;

import exceltoexcel.Sample;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class Writer {
    //эта переменная указывает, с какой строки начнётся запись в файл
    int currentRow = 7;

    public void writeToExcel(HashMap<String, String> numbersOfCellsToWrite, ArrayList<Sample> samplesToWrite) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheetToWrite = workbook.createSheet("HallOffice2023TestSheet");
        for (Sample sampleToWrite : samplesToWrite) {
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
