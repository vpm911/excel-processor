package com.vishal;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

public class ExcelProcessor {

    private static final Set<Integer> columnsToCheck = new HashSet<>();
    private static final int COL_1 = 1;

    static {
        columnsToCheck.add(1);
        columnsToCheck.add(3);
        columnsToCheck.add(4);
        columnsToCheck.add(5);
        columnsToCheck.add(7);
        columnsToCheck.add(8);
        columnsToCheck.add(9);
        columnsToCheck.add(10);
        columnsToCheck.add(11);
        columnsToCheck.add(12);
        columnsToCheck.add(13);
        columnsToCheck.add(14);
        columnsToCheck.add(15);

    }

    public static void main(String[] args) throws Exception {
        String inputFilePath = args[0], outputFilePath = args[1];


        FileInputStream inputFileStream = new FileInputStream(inputFilePath);
        FileOutputStream outputStream = new FileOutputStream(outputFilePath);
        Workbook outputBook = new Workbook(outputStream, "Vishal", "01.1010");

        try (ReadableWorkbook wb = new ReadableWorkbook(inputFileStream)) {
            Sheet iSheet = wb.getFirstSheet();
            Worksheet oSheet = outputBook.newWorksheet(iSheet.getName());

            try (Stream<Row> rows = iSheet.openStream()) {
                int r = 0;
                List<Row> rowList = rows.toList();
                for (Row row : rowList) {
                    String propertyCode = row.getCell(COL_1).getText();
                    if (propertyCode.contains("/")) {
                        String[] split = propertyCode.split("/");
                        for (int i = 0; i < split.length; i++, r++) {
                            for (int c = 0; c < row.getCellCount(); c++) {
                                if (columnsToCheck.contains(c) && row.getCell(c).getText().contains("/")) {
                                    oSheet.value(r, c, row.getCell(c).getText().split("/")[i]);
                                } else {
                                    oSheet.value(r, c, row.getCell(c).getRawValue());
                                }
                            }
                        }
                    } else {
                        for (int c = 0; c < row.getCellCount(); c++) {
                            oSheet.value(r, c, row.getCell(c).getRawValue());
                        }
                        r++;
                    }
                }
            }

            oSheet.finish();
            outputBook.finish();
            outputStream.close();
            inputFileStream.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }


          /*  for (List<Object> row : inputList) {
                // if property code contains a /
                if (row.get(COL_1).toString().contains("/")) {
                    List<List<Object>> splitRows = new ArrayList<>();
                    int splitCount = row.get(COL_1).toString().split("/").length;
                    for (int k = 0; k < splitCount; k++) {
                        splitRows.add(new ArrayList<>());
                    }
                    for (int j = 0; j < row.size(); j++) {
                        String content = row.get(j).toString();

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
            } */


    }

}