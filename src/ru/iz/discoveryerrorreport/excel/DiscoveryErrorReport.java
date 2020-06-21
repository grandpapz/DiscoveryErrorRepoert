package ru.iz.discoveryerrorreport.excel;
// Created by Ivan Zasukhin 20/06/2020

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.swing.text.Style;
import java.io.*;

public class DiscoveryErrorReport {
    public static void main(String[] args) throws IOException {
        String file1 = "/Users/ivanzasukhin/Downloads/errors_example/sostav.xls";
        String file2 = "/Users/ivanzasukhin/Downloads/errors_example/errors.xls";
//        Scanner scanner = new Scanner(System.in);
//        System.out.println("Введите полный путь к cоставу среды:");
////        file1 = scanner.nextLine();
//        System.out.println("Введите полный путь к отчету ошибок дискаверинга:");
////        file2 = scanner.nextLine();
//        scanner.close();



        // Open files with Structure and Discovery Errors
        FileInputStream structure = new FileInputStream(new File(file1));
        FileInputStream errors = new FileInputStream(new File(file2));
        Workbook errorBook = new HSSFWorkbook(errors);
        Workbook structBook = new HSSFWorkbook(structure);
        Workbook reportBook = new HSSFWorkbook();
        Sheet eSheet = errorBook.getSheetAt(0);
        Sheet sSheet = structBook.getSheetAt(0);
        int lastRowE = eSheet.getLastRowNum();
        int lastRowS = sSheet.getLastRowNum();

        // Headers style & font
//        CellStyle header = reportBook.createCellStyle();
//        header.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
//        Font headerFont = reportBook.createFont();
//        headerFont.setFontHeight((short) 16);
//        headerFont.setBold(true);
//        header.setAlignment(HorizontalAlignment.forInt(3));
//        header.setFont(headerFont);

        // Create report and set headers at row 0

        Sheet rSheet = reportBook.createSheet("Discovery Errors Report");

        Row row = rSheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Server name:");
        cell = row.createCell(1);
        cell.setCellValue("IP address:");
        cell = row.createCell(2);
        cell.setCellValue("Discovery Status:");
        rSheet.autoSizeColumn(0);
        rSheet.autoSizeColumn(1);
        rSheet.autoSizeColumn(2);
        // Search for errors, write high priority error to report.xls
        for (int i = 1; i <= lastRowS; i++) {
            row = rSheet.createRow(rSheet.getLastRowNum() + 1);
            cell = row.createCell(0);
            cell.setCellValue(sSheet.getRow(i).getCell(0).getStringCellValue());
            cell = row.createCell(1);
            cell.setCellValue(sSheet.getRow(i).getCell(10).getStringCellValue());
            String line = cell.getStringCellValue();

            for (int j = 1; j <= lastRowE; j++) {
                if (line.contains(eSheet.getRow(j).getCell(9).getStringCellValue())) {
                    String error = eSheet.getRow(j).getCell(2).getStringCellValue();
                    for (int r = 1; r < lastRowE; r++) {
                        if (eSheet.getRow(r).getCell(9).getStringCellValue().equals(eSheet.getRow(j).getCell(9).getStringCellValue())) {
                            if (errorPriority(eSheet.getRow(r).getCell(2).getStringCellValue()) > errorPriority(eSheet.getRow(j).getCell(2).getStringCellValue())) {
                                error = eSheet.getRow(r).getCell(2).getStringCellValue();
                            }
                        }
                    }
                    cell = row.createCell(2);
                    cell.setCellValue(error);
                }
            }
        }
        structBook.close();
        errorBook.close();
        FileOutputStream report = new FileOutputStream(new File("/Users/ivanzasukhin/Desktop/report.xls"));
        reportBook.write(report);
        reportBook.close();

    }

    private static int errorPriority(String error) {
        int priority = 0;
        if (error.contains("NTCMD: Internal error. Details: Server not reachable by netbios") ||
                error.contains("SSH: Permission denied.") || error.contains("NTCMD: Network path is not accessible")){
            priority = 2;
        }
        else if(error.equals("NTCMD: Invalid user name or password.") ||
                error.equals("SSH: Disconnecting because key exchange failed at the local or remote end.") ||
                error.equals("SSH: Internal error. Details: Key exchange failed: server's signature didn't verify (uc)")){
            priority = 1;
        }

        return priority;
    }
}
