package com.excelreader;

import com.sun.speech.freetts.Voice;
import com.sun.speech.freetts.VoiceManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;

public class ExcelReader {
    private Voice voice;
    private boolean isReading;
    private String currentFile;
    private int currentRow = -1;

    public ExcelReader() {
        System.setProperty("freetts.voices", "com.sun.speech.freetts.en.us.cmu_us_kal.KevinVoiceDirectory");
        voice = VoiceManager.getInstance().getVoice("kevin16");
        if (voice != null) {
            voice.allocate();
        }
    }

    public void startReading(String column) {
        isReading = true;
        try {
            if (currentFile == null) return;
            
            FileInputStream fis = new FileInputStream(new File(currentFile));
            Workbook workbook;
            
            if (currentFile.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else {
                workbook = new HSSFWorkbook(fis);
            }

            Sheet sheet = workbook.getSheetAt(0);
            int columnIndex = convertColumnToIndex(column);
            
            if (currentRow == -1) {
                currentRow = sheet.getFirstRowNum();
            }

            Row row = sheet.getRow(currentRow);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    String cellValue = getCellValue(cell);
                    if (voice != null && cellValue != null) {
                        voice.speak(cellValue);
                    }
                }
            }

            workbook.close();
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void setCurrentFile(String file) {
        this.currentFile = file;
    }

    public void setCurrentRow(int row) {
        this.currentRow = row;
    }

    private int convertColumnToIndex(String column) {
        int result = 0;
        for (int i = 0; i < column.length(); i++) {
            result *= 26;
            result += column.charAt(i) - 'A' + 1;
        }
        return result - 1;
    }

    private String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    public void stopReading() {
        isReading = false;
        if (voice != null) {
            voice.deallocate();
        }
    }
} 