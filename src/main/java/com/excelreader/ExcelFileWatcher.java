package com.excelreader;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.Timer;
import java.util.TimerTask;

public class ExcelFileWatcher {
    private Timer timer;
    private Consumer<List<String>> onExcelFilesChanged;

    public void setOnExcelFilesChanged(Consumer<List<String>> handler) {
        this.onExcelFilesChanged = handler;
    }

    public void start() {
        timer = new Timer();
        timer.scheduleAtFixedRate(new TimerTask() {
            @Override
            public void run() {
                List<String> excelFiles = findOpenExcelFiles();
                if (onExcelFilesChanged != null) {
                    onExcelFilesChanged.accept(excelFiles);
                }
            }
        }, 0, 1000);
    }

    public void stop() {
        if (timer != null) {
            timer.cancel();
            timer = null;
        }
    }

    private List<String> findOpenExcelFiles() {
        List<String> files = new ArrayList<>();
        // 这里需要根据操作系统实现查找打开的Excel文件的逻辑
        // Windows可以使用JNA访问Windows API
        // Mac可以使用AppleScript
        return files;
    }
} 