package exceltoexcel.serviceClasses;

import exceltoexcel.Sample;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Pattern;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class Extractor {
    ArrayList<Sample> extractedSamples = new ArrayList<>();
    HashMap<String, String> sampleParts = new HashMap<>();
    // в множество unknownStrings собираются те строки, которые не попадают ни под одно из условий экстракции,
    //эти строки могут быть полезны при переналадке программы под другие продукты.
    HashSet<String> unknownStrings = new HashSet<>();

    public XSSFSheet prepareSheetToRead() throws IOException {
        //получаем доступ к листу
        FileInputStream file = new FileInputStream(new File("C:\\Users\\kekec\\Desktop\\OasisHallandOfficeADasha.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        return sheet;
    }

    public ArrayList<Sample> extract(HashMap<String, String> numbersOfAnalyses,
                                     HashMap<String, String> typesOfAnalysis, XSSFSheet sheet) {
        int sampleCounter = 0;
        //цикл, который пройдётся по всему листу
        for (int rowCounterGlobal = 4; rowCounterGlobal < 18700;
             rowCounterGlobal++) {
            //если текущая строка пустая, она пропускается
            Row currentRow = sheet.getRow(rowCounterGlobal);
            if (currentRow == null) {
                continue;
            }
            //получаем первую ячейку в текущей строке
            Cell currentCell = currentRow.getCell(0, CREATE_NULL_AS_BLANK);
            currentCell.setCellType(CellType.STRING);
            String cellContain = currentCell.getStringCellValue();


            //этот сегмент помогает найти неучтённые анализы при переналадке программы на другой продукт. Он ломает
            //выгрузку инженера, поэтому должен закомменчиваться в работе
            if (!typesOfAnalysis.containsKey(cellContain)) {
                //условие отсекает строки, содержащие слэш и кириллицу
                if (!cellContain.contains("/") & !Pattern.matches(".*\\p{InCyrillic}.*", cellContain)) {
                    unknownStrings.add(cellContain);
                }
                continue;
            }
            // этот сегмент проставляет точки между днём, месяцем и годом в датах
            if (cellContain.equals("Custom./Chem.")) {
                StringBuilder dateBuilder = new StringBuilder();
                dateBuilder = dateBuilder.append(currentRow.getCell(3).getStringCellValue())
                        .insert(2, ".").insert(5, ".");
                String date = dateBuilder.toString();
                sampleParts.put("Custom./Chem.", date);
                continue;
            }
            //эта условие считывает положение конца сэмпла и передаёт данные на Writer, затем считывает сэмпл и танк
            if (cellContain.equals("Sample")) {
                HashMap<String, String> sampleParts2 = new HashMap<>();
                sampleParts2.putAll(sampleParts);
                Sample sample = new Sample();
                sample.setSampleParts(sampleParts2);
                System.out.println(sampleParts2.toString());
                extractedSamples.add(sampleCounter, sample);
                sampleParts.clear();
                sampleCounter++;
                performSampleAndTankExtraction(currentRow, sampleParts);
                continue;
            }

            //эта секция проверяет, имеется ли в виду приёмо-сдаточный вид плёнки, либо периодический
            if (cellContain.equals("Film")) {
                performFilmExtraction(currentRow, sampleParts);
            }
            //эта секция вытаскивает бокс
            if (cellContain.equals("Free")) {
                performBoxExtraction(sampleParts, sheet, rowCounterGlobal);
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

            //эта секция проверяет все остальные анализы
            if (typesOfAnalysis.containsKey(cellContain)) {
                performAnalysesExtraction(currentRow, sampleParts, numbersOfAnalyses);
                continue;
            }

        }
        for (Sample sample : extractedSamples) {
            System.out.println(sample.getSampleParts().toString());
        }
        System.out.println(unknownStrings.toString());
        return extractedSamples;
    }

    public void performEngineerExtraction(Row row, HashMap<String, String> sampleParts) {
        sampleParts.put("Responsible", row.getCell(1).getStringCellValue());
    }

    public void performBatchExtraction(Row row, HashMap<String, String> sampleParts) {
        sampleParts.put("Product", row.getCell(4).getStringCellValue());
    }

    public void performFilmExtraction(Row row, HashMap<String, String> sampleParts) {
        if (row.getCell(1).getStringCellValue().equals("color")) {
            sampleParts.put("FilmColor", row.getCell(2).getStringCellValue());
        }
        if (row.getCell(1).getStringCellValue().equals("appearance")) {
            sampleParts.put("FilmAppearance", row.getCell(2).getStringCellValue());
        }
    }

    public void performSampleAndTankExtraction(Row row, HashMap<String, String> sampleParts) {
        String sample = row.getCell(2).getStringCellValue();
        Character lastSymbolOfSample = sample.charAt(sample.length() - 1);
        String lastNumberOfSample = lastSymbolOfSample.toString();
        sampleParts.put("Sample", lastNumberOfSample);
        //эта проверка нужна потому, что инженеры НЕ ВСЕГДА забивают танк, и на пустой ячейке программа крашится
        if (row.getCell(4) != null) {
            sampleParts.put("Tank", row.getCell(4).getStringCellValue());
        }
    }

    public void performCommentExtraction(HashMap<String, String> sampleParts, XSSFSheet sheet, int rowCounterGlobal) {
        StringBuilder comment = new StringBuilder();
        int rowCounter = 0;

        //этот вложенный цикл размером 12на25 объединяет ячейки после ключевого слова "Comment".
        while (rowCounter < 12) {
            Row row = sheet.getRow(rowCounterGlobal + rowCounter + 1);
            int cellCounter = 0;
            if (row == null) {
                rowCounter++;
                continue;
            }
            while (cellCounter < 25) {
                Cell cell = row.getCell(cellCounter, CREATE_NULL_AS_BLANK);
                if (cell.getCellType() == CellType.NUMERIC) {
                    cell.setCellType(CellType.STRING);
                }
                comment.append(cell.getStringCellValue() + " ");
                cellCounter++;
            }
            rowCounter++;
        }
        //эта секция фильтрует лишние части комментария
        String comment2 = StringUtils.substringBefore(String.valueOf(comment), "....");
        String readyComment = StringUtils.substringBefore(String.valueOf(comment2), "Comments");
        sampleParts.put("Comments", readyComment);
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
        }
    }

    public void performAnalysesExtraction(Row row, HashMap<String, String> sampleParts,
                                          HashMap<String, String> numbersOfAnalyses) {
        Cell currentCell = row.getCell(0);
        Cell cellWithValueOfAnalysis = row.getCell(Integer.parseInt(numbersOfAnalyses.get(currentCell.getStringCellValue())));
        cellWithValueOfAnalysis.setCellType(CellType.STRING);
        currentCell.setCellType(CellType.STRING);
        sampleParts.put(currentCell.getStringCellValue(), cellWithValueOfAnalysis.getStringCellValue());
    }
}
