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

    public static void createExcel() {
        // Define the sheet name
        String sheetName = "Data";

        // Create a new workbook and select the sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);

        // Define headers
        Row headers = sheet.createRow(0);
        Cell cell1 = headers.createCell(0);
        cell1.setCellValue("Jira Team");

        Cell cell2 = headers.createCell(1);
        cell2.setCellValue("Automation Count");

        Cell cell3 = headers.createCell(2);
        cell3.setCellValue("Manual Count");

        // Add data rows
        Row dataRow1 = sheet.createRow(1);
        Cell cell4 = dataRow1.createCell(0);
        cell4.setCellValue("Team A");

        Cell cell5 = dataRow1.createCell(1);
        cell5.setCellValue(10);

        Cell cell6 = dataRow1.createCell(2);
        cell6.setCellValue(5);

        Row dataRow2 = sheet.createRow(2);
        Cell cell7 = dataRow2.createCell(0);
        cell7.setCellValue("Team B");

        Cell cell8 = dataRow2.createCell(1);
        cell8.setCellValue(5);

        Cell cell9 = dataRow2.createCell(2);
        cell9.setCellValue(15);

        // Write the workbook to a file
        FileOutputStream fileOut = new FileOutputStream("data.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        workbook.close();
        }
    }
}

