import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Application {

    public static void main(String[] args) throws Exception {
        Application a = new Application();
        //a.processExcel();
        a.processExcelNew();
    }

    private void processExcelSame() throws Exception {

        FileInputStream fileInputStream = new FileInputStream(new File("/Users/Z003985/repos/work-with-excel/src/main/resources/input_test.xlsx"));
        Workbook workbook_input = new XSSFWorkbook(fileInputStream);

        // Create an output Excel to add the cell values from the input
        Workbook workbook_output = new XSSFWorkbook();

        // Reading the Excel
        int totalSheets = workbook_input.getNumberOfSheets();
        System.out.println("Number of sheets: " + totalSheets);
        for (int sheetNumber = 0; sheetNumber < totalSheets; sheetNumber++) {
            Sheet sheet = workbook_input.getSheetAt(sheetNumber);
            sheet.getSheetConditionalFormatting();
            String sheetName = sheet.getSheetName();
            System.out.println("Extracting data for " + sheetNumber + ". Name of sheet: " + sheetName);

            List<List<String>> rows = new ArrayList<>();

            int ithRow = 0;
            for (Row row : sheet) { // Adding all rows
                rows.add(new ArrayList());
                for (Cell cell : row) { // Adding all columns
                    // Assuming all data is in string format only
                    rows.get(ithRow).add(cell.getRichStringCellValue().getString());
                }
                ithRow++;
            }
            //workbook_input.close();
            fileInputStream.close();

            System.out.println("Read all formatted and colored headers as simple text");
            System.out.println("Rows: " + rows.size());
            System.out.println("Columns: " + rows.get(0).size());
            for (int row = 0; row < rows.size(); row++) { // Assuming we need to copy the first 22 rows only
                for (int col = 0; col < rows.get(0).size(); col++) {
                    System.out.print(rows.get(row).get(col));
                }
            }

            //Sheet sheet_output = workbook_output.createSheet(sheetName); // Creating a similarly named sheet

            for (int row = 0; row < rows.size(); row++) {
                Row r = sheet.createRow(row);
                for (int col = 0; col < 3; col++) { // Set the columns you want to add only
                    Cell cell = r.createCell(col);
                    cell.setCellValue(rows.get(row).get(col));
                }
            }
            //int lastRow = sheet.getLastRowNum();
            //sheet_output.setAutoFilter(new CellRangeAddress(0, 0, 0, 0));
        }

        // Write to Excel
        FileOutputStream fileOutputStream = new FileOutputStream("/Users/Z003985/repos/work-with-excel/src/main/resources/output_test.xlsx");
        workbook_input.write(fileOutputStream);
        workbook_input.close();
        fileOutputStream.close();
    }

    private void processExcel() throws Exception {

        FileInputStream fileInputStream = new FileInputStream(new File("/Users/Z003985/repos/work-with-excel/src/main/resources/input_test.xlsx"));
        Workbook workbook_input = new XSSFWorkbook(fileInputStream);

        // Create an output Excel to add the cell values from the input
        Workbook workbook_output = new XSSFWorkbook();

        // Reading the Excel
        int totalSheets = workbook_input.getNumberOfSheets();
        System.out.println("Number of sheets: " + totalSheets);
        for (int sheetNumber = 0; sheetNumber < totalSheets; sheetNumber++) {
            Sheet sheet = workbook_input.getSheetAt(sheetNumber);
            String sheetName = sheet.getSheetName();
            System.out.println("Extracting data for " + sheetNumber + ". Name of sheet: " + sheetName);

            List<List<String>> rows = new ArrayList<>();

            int ithRow = 0;
            for (Row row : sheet) { // Adding all rows
                rows.add(new ArrayList());
                for (Cell cell : row) { // Adding all columns
                    // Assuming all data is in string format only
                    rows.get(ithRow).add(cell.getRichStringCellValue().getString());
                }
                ithRow++;
            }
            workbook_input.close();
            fileInputStream.close();

            System.out.println("Read all formatted and colored headers as simple text");
            System.out.println("Rows: " + rows.size());
            System.out.println("Columns: " + rows.get(0).size());
            for (int row = 0; row < rows.size(); row++) { // Assuming we need to copy the first 22 rows only
                for (int col = 0; col < rows.get(0).size(); col++) {
                    System.out.print(rows.get(row).get(col));
                }
            }

            Sheet sheet_output = workbook_output.createSheet(sheetName); // Creating a similarly named sheet

            for (int row = 0; row < rows.size(); row++) {
                Row r = sheet_output.createRow(row);
                for (int col = 0; col < 3; col++) { // Set the columns you want to add only
                    Cell cell = r.createCell(col);
                    cell.setCellValue(rows.get(row).get(col));
                }
            }
        }

        // Write to Excel
        FileOutputStream fileOutputStream = new FileOutputStream("/Users/Z003985/repos/work-with-excel/src/main/resources/output_test.xlsx");
        workbook_output.write(fileOutputStream);
        workbook_output.close();
        fileOutputStream.close();
    }

    private void processExcelNew() throws Exception {

        FileInputStream fileInputStream = new FileInputStream(new File("/Users/Z003985/repos/work-with-excel/src/main/resources/input_test.xlsx"));
        Workbook workbook_input = new XSSFWorkbook(fileInputStream);

        // Create an output Excel to add the cell values from the input
        //Workbook workbook_output = new XSSFWorkbook();

        // Reading the Excel
        int totalSheets = workbook_input.getNumberOfSheets();
        System.out.println("Number of sheets: " + totalSheets);
        for (int sheetNumber = 0; sheetNumber < totalSheets; sheetNumber++) {
            Sheet sheet = workbook_input.getSheetAt(sheetNumber);
            String sheetName = sheet.getSheetName();
            System.out.println("Extracting data for " + sheetNumber + ". Name of sheet: " + sheetName);

            List<List<String>> rows = new ArrayList<>();

            int ithRow = 0;
            for (Row row : sheet) { // Adding all rows
                rows.add(new ArrayList());
                for (Cell cell : row) { // Adding all columns
                    switch (cell.getCellType()) {
                        case STRING:
                            rows.get(ithRow).add(cell.getRichStringCellValue().getString());
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                rows.get(ithRow).add(cell.getDateCellValue() + "");
                            } else {
                                rows.get(ithRow).add(cell.getNumericCellValue() + "");
                            }
                            break;
                        case BOOLEAN:
                            rows.get(ithRow).add(cell.getBooleanCellValue() + "");
                            break;
                        case FORMULA:
                            rows.get(ithRow).add(cell.getCellFormula() + "");
                            break;
                        default:
                            throw new Exception("Undefined type in Sheet (" + sheetNumber + ") " + sheetName);
                    }
                }
                ithRow++;
            }
            //workbook_input.close();
            //fileInputStream.close();

            System.out.println("Read all formatted and colored headers as simple text");
            System.out.println("Rows: " + rows.size());
            System.out.println("Columns: " + rows.get(0).size());
            for (int row = 0; row < rows.size(); row++) { // Assuming we need to copy the first 22 rows only
                for (int col = 0; col < rows.get(0).size(); col++) {
                    System.out.print(rows.get(row).get(col));
                }
            }

            workbook_input.removeSheetAt(sheetNumber); // Remove the current sheet after the string/text data has been extracted in Arraylist
            Sheet sheet_output = workbook_input.createSheet(sheetName); // Adding similarly named sheet at the end
            workbook_input.setSheetOrder(sheetName, sheetNumber);

            for (int row = 0; row < rows.size(); row++) {
                Row r = sheet_output.createRow(row);
                for (int col = 0; col < 3; col++) { // Set the columns you want to add only
                    Cell cell = r.createCell(col);
                    cell.setCellValue(rows.get(row).get(col));
                }
            }
        }

        // Write to Excel
        FileOutputStream fileOutputStream = new FileOutputStream("/Users/Z003985/repos/work-with-excel/src/main/resources/input_test.xlsx");
        //workbook_output.write(fileOutputStream);
        //workbook_output.close();
        //fileOutputStream.close();
        workbook_input.write(fileOutputStream);
        workbook_input.close();
        fileOutputStream.close();
    }
}
