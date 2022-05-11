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

        FileInputStream fileInputStream = new FileInputStream(new File("/<path>/input_test.xlsx"));
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        // Reading the Excel
        Sheet sheet = workbook.getSheetAt(0);
        List<String> headers = new ArrayList<String>();

        for (Cell cell : sheet.getRow(0)) { // Adding the headers from the first row only
            switch (cell.getCellType()) {
                case STRING:
                    headers.add(cell.getRichStringCellValue().getString());
                    break;
                default:
                    throw new Exception("Not string header");
            }
        }

        fileInputStream.close();

        System.out.println("Read all formatted and colored headers as simple text");
        for (int i = 1; i <= headers.size(); i++) {
            System.out.println(i + ". " + headers.get(i-1));
        }

        // Overriding the Excel header cells
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }

        // Write to Excel
        FileOutputStream fileOutputStream = new FileOutputStream("/<path>/output_test.xlsx");
        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();
    }
}
