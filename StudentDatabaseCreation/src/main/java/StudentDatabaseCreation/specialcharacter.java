package StudentDatabaseCreation;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;

public class specialcharacter {
    public static void main(String[] args) {
        String filePath = "D:/StudentDatabase.xlsx";

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming you want to check the first sheet

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        if (containsSpecialCharacters(cellValue)) {
                            System.out.println("Special characters found in cell: " + cellValue);
                        }
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean containsSpecialCharacters(String value) {
        String specialCharacters = "!@#$%^&*()_+{}[]|\';\":;/<>?,.";
        for (char c : value.toCharArray()) {
            if (specialCharacters.contains(Character.toString(c))) {
                return true;
            }
        }
        return false;
    }
}
