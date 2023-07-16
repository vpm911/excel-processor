package com.vishal;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.Cell;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.stream.Stream;

public class ExcelProcessor {

    private static final Set<Integer> columnsToCheck = new HashSet<>();
    private static final int COL_1 = 1;

    static {
        columnsToCheck.add(1);
        columnsToCheck.add(3);
        columnsToCheck.add(4);
        columnsToCheck.add(5);
    }

    public static void main(String[] args) throws Exception {
        String inputFilePath = args[0], outputFilePath = args[1];
        List<List<String>> inputList = new ArrayList<>();
        List<List<String>> outputList = new ArrayList<>();

        FileInputStream inputFileStream = new FileInputStream(inputFilePath);
        FileOutputStream outputStream = new FileOutputStream(outputFilePath);
        Workbook outputBook = new Workbook(outputStream, "Vishal", "01.1010");

        try (ReadableWorkbook wb = new ReadableWorkbook(inputFileStream)) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                rows.forEach(row -> {
                    List<String> rowData = row.getCells(0, row.getCellCount()).stream().map(Cell::getText).toList();
                    List<Cell> rowData1 = row.getCells(0, row.getCellCount()).stream().toList();
                    inputList.add(rowData);
                });
                inputFileStream.close();
            } catch (Exception ex) {
                ex.printStackTrace();
            }

            for (List<String> row : inputList) {
                // if property code contains a /
                if (row.get(COL_1).contains("/")) {
                    List<List<String>> splitRows = new ArrayList<>();
                    int splitCount = row.get(COL_1).split("/").length;
                    for (int k = 0; k < splitCount; k++) {
                        splitRows.add(new ArrayList<>());
                    }
                    for (int j = 0; j < row.size(); j++) {
                        String content = row.get(j);

                        // if column data is supposed to be split
                        if (columnsToCheck.contains(j) && content.contains("/")) {
                            String[] split = content.split("/");
                            for (int k = 0; k < splitCount; k++) {
                                splitRows.get(k).add(split[k]);
                            }
                        } else {
                            for (int k = 0; k < splitCount; k++) {
                                splitRows.get(k).add(content);
                            }
                        }
                    }
                    outputList.addAll(splitRows);
                } else {
                    outputList.add(row);
                }
            }


            Worksheet outSheet = outputBook.newWorksheet(sheet.getName());
            for (int i = 0; i < outputList.size(); i++) {
                for (int j = 0; j < outputList.get(i).size(); j++) {
                    outSheet.value(i, j, outputList.get(i).get(j));
                }
            }

            outputBook.finish();
        }

    }

}