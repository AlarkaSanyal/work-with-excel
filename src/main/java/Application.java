import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Application {

    public static void main(String[] args) throws Exception {
        Application a = new Application();
        a.processExcel();
    }

    private void processExcel() throws Exception {

        FileInputStream fileInputStream = new FileInputStream(new File("/Users/Z003985/repos/work-with-excel/src/main/resources/input_test.xlsx"));
        Workbook workbook_input = new XSSFWorkbook(fileInputStream);

        // Reading the Excel
        Sheet sheet = workbook_input.getSheetAt(0); // First sheet only

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

        // Create an output Excel and add the cell values from the ArrayList
        Workbook workbook_output = new XSSFWorkbook();
        Sheet sheet_output = workbook_output.createSheet();
        for (int row = 0; row < rows.size(); row++) {
            Row r = sheet_output.createRow(row);
            for (int col = 0; col < 3; col++) { // Set the columns you want to add only
                Cell cell = r.createCell(col);
                cell.setCellValue(rows.get(row).get(col));
            }
        }

        // Write to Excel
        FileOutputStream fileOutputStream = new FileOutputStream("/Users/Z003985/repos/work-with-excel/src/main/resources/output_test.xlsx");
        workbook_output.write(fileOutputStream);
        workbook_output.close();
        fileOutputStream.close();
    }
}
