package exceltoexcel.serviceClasses;

import exceltoexcel.Sample;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class Extractor {


    public XSSFSheet prepareSheet() throws IOException {
        //получаем доступ к листу эксель
        FileInputStream file = new FileInputStream(new File("C:\\Users\\kekec\\Desktop\\OasisHallandOfficeADasha.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        return sheet;


    }

    public HashMap<String, String> prepareTypesOfAnalyses() throws IOException {

        //получаем доступ к файлу пропертис с именами анализов
        String rootPath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
        String appConfigPath = rootPath + "analyses_reading_names.properties";
        Properties readingProperties = new Properties();
        readingProperties.load(new FileInputStream(appConfigPath));
        HashMap<String, String> typesOfAnalyses = new HashMap<>();
        //считываем все записи из пропертис и заполняем хэшмапу с типами испытаний
        for (Map.Entry<Object, Object> propertiesEntrySet : readingProperties.entrySet()) {
            typesOfAnalyses.put((String) propertiesEntrySet.getKey(), (String) propertiesEntrySet.getValue());
        }
        // System.out.println(typesOfAnalyses.toString());
        return typesOfAnalyses;
    }

    public HashMap<String, String> prepareNumbersOfCellWithAnalyse() throws IOException {

        //получаем доступ к файлу пропертис с именами анализов
        String rootPath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
        String appConfigPath = rootPath + "analyses_reading_numbers_of_cells.properties";
        Properties readingProperties = new Properties();
        readingProperties.load(new FileInputStream(appConfigPath));
        HashMap<String, String> numbersOfAnalyses = new HashMap<>();
        //считываем все записи из пропертис и заполняем хэшмапу с типами испытаний
        for (Map.Entry<Object, Object> propertiesEntrySet : readingProperties.entrySet()) {
            numbersOfAnalyses.put((String) propertiesEntrySet.getKey(), (String) propertiesEntrySet.getValue());
        }
        // System.out.println(typesOfAnalyses.toString());
        return numbersOfAnalyses;
    }

    public ArrayList<Sample> extract(HashMap<String, String> numbersOfAnalyses,
                                     HashMap<String, String> typesOfAnalysis, XSSFSheet sheet) {
        HashMap<String, String> sampleParts = new HashMap<>();
        ArrayList<Sample> extractedSamples = new ArrayList<>();
        int sampleCounter = 0;
        //цикл, который пройдётся по всему листу
        for (int rowCounterGlobal = 0; rowCounterGlobal < Integer.parseInt(typesOfAnalysis.get("RowsToExtract"));
             rowCounterGlobal++) {

            //если текущая строка пустая, она пропускается
            Row currentRow = sheet.getRow(rowCounterGlobal);
            if (currentRow == null) {
                continue;
            }
            //получаем первую ячейку в текущей строке
            String cellContain = currentRow.getCell(0, CREATE_NULL_AS_BLANK).getStringCellValue();
            //если ячейка пустая - пропускаем
       //     if (cellContain == null) {
       //         continue;
       //     }
            //это условие отбрасывает все строки, которые не удовлетворяют условиям поиска
            if (!typesOfAnalysis.containsKey(cellContain)) {
                continue;
            }
            //эта условие считывает положение конца сэмпла и передаёт данные на Writer
            if (Objects.equals(cellContain, "S")) {
                sampleCounter++;
             //   System.out.println(sampleParts);
                Sample sample = new Sample(sampleParts);
                extractedSamples.add(sample);
                sampleParts.clear();

                continue;
            }

            //эта секция вытаскивает сэмпл и танк
            if (typesOfAnalysis.get(cellContain).equals("Sample")) {
                performSampleAndTankExtraction(currentRow, sampleParts);
                continue;
            }
            //эта секция вытаскивает коммент
            if (typesOfAnalysis.get(cellContain).equals("Comments")) {
                //этот участок служит для того, чтобы не брать в расчёт комментарии к анализам
                if (currentRow.getCell(2, CREATE_NULL_AS_BLANK).getStringCellValue().equals("analysis")) {
                    continue;
                }
                performCommentExtraction(sampleParts, sheet, rowCounterGlobal);
                continue;
            }
            //эта секция вытаскивает бокс
            if (typesOfAnalysis.get(cellContain).equals("Free")) {
                performBoxExtraction(sampleParts, sheet, rowCounterGlobal);
                continue;
            }
            //эта секция проверяет все остальные анализы
            if (typesOfAnalysis.containsKey(cellContain)) {
                performAnalysesExtraction(currentRow, sampleParts, numbersOfAnalyses);
                continue;
            }



        }
        return extractedSamples;
    }







    public void performSampleAndTankExtraction(Row row, HashMap<String, String> sampleParts) {
        sampleParts.put("Sample", row.getCell(2).getStringCellValue());
        //эта проверка нужна потому, что инженеры НЕ ВСЕГДА забивают танк, и на пустой ячейке программа крашится
        if (row.getCell(4) != null) {
            sampleParts.put("Tank", row.getCell(4).getStringCellValue());
        }
    }

    public void performCommentExtraction(HashMap<String, String> sampleParts, XSSFSheet sheet, int rowCounterGlobal) {
        StringBuilder comment = new StringBuilder();
        int rowCounter = 0;

        //этот вложенный цикл размером 7на10 объединяет 70 ячеек после ключевого слова "Comment".
        while (rowCounter < 10) {
            Row row = sheet.getRow(rowCounterGlobal + rowCounter + 1);
            int cellCounter = 0;
            if (row == null) {
                rowCounter++;
                continue;
            }
            while (cellCounter < 7) {
                Cell cell = row.getCell(cellCounter, CREATE_NULL_AS_BLANK);
                comment.append(cell.getStringCellValue() + " ");
                cellCounter++;
            }
            rowCounter++;
        }
        //эта хуёво написанная секция фильтрует лишние части комментария. Работает, значит, и переделывать не нужно
        String comment2 = StringUtils.substringBefore(String.valueOf(comment), "....");
        String readyComment = StringUtils.substringBefore(String.valueOf(comment2), "Comments");
        sampleParts.put("Comments", readyComment);
        //   System.out.println(sampleParts.toString());
    }

    public void performBoxExtraction(HashMap<String, String> sampleParts, XSSFSheet sheet, int rowCounterGlobal) {
        StringBuilder freeText = new StringBuilder();
        int rowCounter = 0;

        //этот вложенный цикл размером 5на5 объединяет 25 ячеек после ключевого слова "Free".
        while (rowCounter < 5) {
            Row row = sheet.getRow(rowCounterGlobal + rowCounter + 1);
            int cellCounter = 0;
            if (row == null) {
                rowCounter++;
                continue;
            }
            while (cellCounter < 5) {
                Cell cell = row.getCell(cellCounter, CREATE_NULL_AS_BLANK);
                freeText.append(cell.getStringCellValue() + " ");
                cellCounter++;
            }
            rowCounter++;
            sampleParts.put("Free", freeText.toString());
            //     System.out.println(sampleParts.toString());
        }
    }

    public void performAnalysesExtraction(Row row, HashMap<String, String> sampleParts,
                                          HashMap<String, String> numbersOfAnalyses) {
        sampleParts.put(row.getCell(0).getStringCellValue(),
                row.getCell(Integer.parseInt(numbersOfAnalyses.get(row.getCell(0).getStringCellValue()))).getStringCellValue());

    }
}
