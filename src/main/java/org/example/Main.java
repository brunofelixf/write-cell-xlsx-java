package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;

public class Main {
    public static void main(String[] args) {

        String arquivoXLSX = "modelo.xlsx";

        try (FileInputStream arquivo = new FileInputStream(new File(arquivoXLSX));
             Workbook workbook = new XSSFWorkbook(arquivo)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(1);
            Cell cell = row.getCell(0);
            System.out.println(cell);

            cell.setCellValue("Novo Valor");

            try (FileOutputStream outputStream = new FileOutputStream(arquivoXLSX)) {
                workbook.write(outputStream);
                System.out.println("O valor da célula A2 foi alterado para 'Novo Valor' com sucesso!");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}