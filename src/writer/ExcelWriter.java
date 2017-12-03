package writer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import processing.Collocation;
import processing.WordStemming;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.List;

public class ExcelWriter {

    public void test() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("FIRST");

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(1);

        // Create a cell and put a value in it.
        Cell cell = row.createCell(1);
        cell.setCellValue(4);

        // Write the output to a file
        FileOutputStream stream = new FileOutputStream("src/result/test.xlsx");
        workbook.write(stream);
        stream.close();
    }

    public void features() throws IOException, ParseException {
        int rowNumber = 1, cellNumber = 1;
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(rowNumber);

        Cell cell = row.createCell(cellNumber);
        cell.setCellValue("WORD");
        cellNumber++;

        cell = row.createCell(cellNumber);
        cell.setCellValue("SynSET");
        cellNumber++;

        cell = row.createCell(cellNumber);
        cell.setCellValue("DEFINITION");
        cellNumber++;

        cell = row.createCell(cellNumber);
        cell.setCellValue("TEXT");

        final byte CODE_TRAINING = 1, CODE_FORBIDDEN = 4, CODE_CONJUNCTION = 5, CODE_FEATURES = 4;
        JSONArray arrayTraining = getArrayData(CODE_TRAINING), arrayForbidden = getArrayData(CODE_FORBIDDEN),
                arrayConjunction = getArrayData(CODE_CONJUNCTION);
        WordStemming stemming = new WordStemming();
        Collocation collocation = new Collocation(arrayConjunction);
        JSONObject data;
        String TEXT;
        String[] splitText;
        List<String> list;
        for (Object o : arrayTraining) {
            data = (JSONObject) o;
            TEXT = (String) data.get("text");
            TEXT = TEXT.replaceAll("\\W", " ");
            splitText = TEXT.split("\\s+");
            for (int i=0; i < splitText.length; i++) {
                if (!isForbidden(splitText[i], arrayForbidden)) {
                    splitText[i] = stemming.getStemming().lemmatize(splitText[i]);
                } else {
                    splitText[i] = splitText[i].toLowerCase();
                }
            }
            list = collocation.process(Arrays.asList(splitText), (String) data.get("POS"));
            for (String s : list) {
                rowNumber++;
                cellNumber--; cellNumber--; cellNumber--;
                row = sheet.createRow(rowNumber); // ADD NEW ROW
                cell = row.createCell(cellNumber);
                cell.setCellValue(s); // SET WORD
                cellNumber++;

                cell = row.createCell(cellNumber);
                cell.setCellValue((String) data.get("sense")); // SET SynSET
                cellNumber++; cellNumber++;

                cell = row.createCell(cellNumber);
                cell.setCellValue(TEXT); // SET TEXT
            }
        }
        if (writeData(CODE_FEATURES, workbook)) {
            System.out.println("WRITE FILE SUCCESS");
        } else {
            System.out.println("ERROR - SOMETHING HAPPEN WHEN WRITING FILE !");
        }
    }

    public void senseTAG() throws IOException, ParseException {
        final byte CODE = 3;
        int rowNumber = 1, cellNumber = 1;
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(rowNumber);

        Cell cell = row.createCell(cellNumber);
        cell.setCellValue("WORD");
        cellNumber++;

        cell = row.createCell(cellNumber);
        cell.setCellValue("SENSE");
        cellNumber++;

        cell = row.createCell(cellNumber);
        cell.setCellValue("DESCRIPTION");

        JSONArray arraySenseTAG = getArrayData(CODE), arraySense;
        JSONObject data, dataSense;
        for (Object o : arraySenseTAG) {
            data = (JSONObject) o;
            arraySense = (JSONArray) data.get("sense");
            for (Object o1 : arraySense) {
                rowNumber++;
                cellNumber--; cellNumber--;
                row = sheet.createRow(rowNumber);

                dataSense = (JSONObject) o1;
                cell = row.createCell(cellNumber);
                cell.setCellValue((String) data.get("POS"));
                cellNumber++;

                cell = row.createCell(cellNumber);
                cell.setCellValue((String) dataSense.get("name"));
                cellNumber++;

                cell = row.createCell(cellNumber);
                cell.setCellValue((String) dataSense.get("description"));
            }
        }

        if (writeData(CODE, workbook)) {
            System.out.println("WRITE FILE SUCCESS");
        } else {
            System.out.println("ERROR - SOMETHING HAPPEN WHEN WRITING FILE !");
        }
    }

    public void testing() throws IOException, ParseException {
        final byte CODE = 2;
        int rowNumber = 1, cellNumber = 1;
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(rowNumber);

        Cell cell = row.createCell(cellNumber);
        cell.setCellValue("WORD");
        cellNumber++;

        cell = row.createCell(cellNumber);
        cell.setCellValue("SENSE");
        cellNumber++;

        cell = row.createCell(cellNumber);
        cell.setCellValue("TEXT");

        JSONArray arrayTesting = getArrayData(CODE), arrayTestcase;
        JSONObject data, dataTestcase;
        for (Object o : arrayTesting) {
            data = (JSONObject) o;
            arrayTestcase = (JSONArray) data.get("testcase");
            for (Object o1 : arrayTestcase) {
                rowNumber++;
                cellNumber--; cellNumber--;
                row = sheet.createRow(rowNumber);

                dataTestcase = (JSONObject) o1;
                cell = row.createCell(cellNumber);
                cell.setCellValue((String) data.get("POS"));
                cellNumber++;

                cell = row.createCell(cellNumber);
                cell.setCellValue((String) dataTestcase.get("sense"));
                cellNumber++;

                cell = row.createCell(cellNumber);
                cell.setCellValue((String) dataTestcase.get("text"));
            }
        }

        if (writeData(CODE, workbook)) {
            System.out.println("WRITE FILE SUCCESS");
        } else {
            System.out.println("ERROR - SOMETHING HAPPEN WHEN WRITING FILE !");
        }
    }

    public void training() throws IOException, ParseException {
        final byte CODE = 1;
        int rowNumber = 1, cellNumber = 1;
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(rowNumber);

        Cell cell = row.createCell(cellNumber);
        cell.setCellValue("WORD");
        cellNumber++;

        cell = row.createCell(cellNumber);
        cell.setCellValue("SENSE");
        cellNumber++;

        cell = row.createCell(cellNumber);
        cell.setCellValue("TEXT");

        JSONArray arrayTraining = getArrayData(CODE);
        JSONObject data;
        for (Object o : arrayTraining) {
            rowNumber++;
            cellNumber--; cellNumber--;
            row = sheet.createRow(rowNumber);

            data = (JSONObject) o;
            cell = row.createCell(cellNumber);
            cell.setCellValue((String) data.get("POS"));
            cellNumber++;

            cell = row.createCell(cellNumber);
            cell.setCellValue((String) data.get("sense"));
            cellNumber++;

            cell = row.createCell(cellNumber);
            cell.setCellValue((String) data.get("text"));
        }

        if (writeData(CODE, workbook)) {
            System.out.println("WRITE FILE SUCCESS");
        } else {
            System.out.println("ERROR - SOMETHING HAPPEN WHEN WRITING FILE !");
        }
    }

    private boolean writeData(byte CODE, Workbook workbook) {
        boolean success = true;
        try {
            String destinationPath;
            switch (CODE) {
                case 1 :
                    destinationPath = "src/result/training.xlsx";
                    break;
                case 2 :
                    destinationPath = "src/result/testing.xlsx";
                    break;
                case 3 :
                    destinationPath = "src/result/senseTAG.xlsx";
                    break;
                case 4 :
                    destinationPath = "src/result/features.xlsx";
                    break;
                default :
                    destinationPath = "src/result/test.xlsx";
                    break;
            }
            FileOutputStream stream = new FileOutputStream(destinationPath);
            workbook.write(stream);
            stream.close();
        } catch (IOException e) {
            success = false;
            e.printStackTrace();
        }
        return success;
    }

    private JSONArray getArrayData(byte CODE) throws IOException, ParseException {
        String resourcePath;
        switch (CODE) {
            case 1 :
                resourcePath = "/resources/training/datatraining.json";
                break;
            case 2 :
                resourcePath = "/resources/testing/datatesting.json";
                break;
            case 3 :
                resourcePath = "/resources/datalibrary/sensetag.json";
                break;
            case 4 :
                resourcePath = "/resources/datalibrary/forbidden.json";
                break;
            case 5 :
                resourcePath = "/resources/datalibrary/conjunction.json";
                break;
            default :
                resourcePath = "/resources/datalibrary/sensetag.json";
                break;
        }
        InputStreamReader reader = new InputStreamReader(getClass().getResourceAsStream(resourcePath));
        JSONParser parser = new JSONParser();
        JSONArray array = (JSONArray) parser.parse(reader);
        reader.close();
        return array;
    }

    private boolean isForbidden(String s, JSONArray arrayForbidden) {
        boolean result = false;
        for (Object o : arrayForbidden) {
            if (o.toString().equalsIgnoreCase(s)) {
                result = true;
                break;
            }
        }
        return result;
    }

}