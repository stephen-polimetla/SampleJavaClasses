import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {

    public static void main(String[] args) {
        String filePath = "C:\\Users\\002U7C744\\Downloads\\Jira Teams.xlsx";
        String teamNames = "";
        try (FileInputStream fileInputStream = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming the teams are in the first sheet
            Row headerRow = sheet.getRow(0);
            Cell teamCell = headerRow.getCell(0); // Assuming the "Jira Teams" column is the first column

            if (teamCell != null && teamCell.getStringCellValue().equals("Jira Teams")) {
                int rowIndex = 1; // Start from the second row

                while (sheet.getRow(rowIndex) != null) {
                    Cell teamNameCell = sheet.getRow(rowIndex).getCell(0);
                    if (teamNameCell != null) {
                        String teamName = "\"" + teamNameCell.getStringCellValue().trim() + "\",";
                        teamNames = teamNames + teamName;
                    }
                    rowIndex++;
                }
            } else {
                System.out.println("Column 'Jira Teams' not found.");
            }
            System.out.println(teamNames.substring(0,teamNames.length()-1));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
