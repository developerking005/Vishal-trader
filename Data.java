import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class ExcelSortAndFilter {
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream("");
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            // Sort the data by date of birth column (assuming it's the first column)
            sheet = sortData(sheet, 0);

            // Filter rows containing "may" in date of birth column
            sheet = filterData(sheet, 0, "may");

            // Write the modified data to a new Excel file
            FileOutputStream outFile = new FileOutputStream("output.xlsx");
            workbook.write(outFile);
            outFile.close();
            workbook.close();
            file.close();

            System.out.println("Excel file processed successfully.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Sheet sortData(Sheet sheet, int columnIndex) {
        DataFormatter formatter = new DataFormatter();
        sheet.sort(Comparator.comparing(row -> {
            Cell cell = row.getCell(columnIndex);
            return formatter.formatCellValue(cell);
        }));
        return sheet;
    }

    private static Sheet filterData(Sheet sheet, int columnIndex, String keyword) {
        for (int i = sheet.getLastRowNum(); i >= 0; i--) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(columnIndex);
            if (!cell.getStringCellValue().toLowerCase().contains(keyword.toLowerCase())) {
                sheet.removeRow(row);
            }
        }
        return sheet;
    }
}
