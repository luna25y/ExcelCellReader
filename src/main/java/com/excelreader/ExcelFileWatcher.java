package com.excelreader;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.Timer;
import java.util.TimerTask;
import com.sun.jna.platform.win32.*;
import com.sun.jna.ptr.PointerByReference;
import com.sun.jna.platform.win32.WinDef.HWND;
import java.io.BufferedReader;
import java.io.InputStreamReader;
import com.sun.jna.Native;

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
        String os = System.getProperty("os.name").toLowerCase();

        if (os.contains("win")) {
            files.addAll(findWindowsExcelFiles());
        } else if (os.contains("mac")) {
            files.addAll(findMacExcelFiles());
        }

        return files;
    }

    private List<String> findWindowsExcelFiles() {
        List<String> files = new ArrayList<>();
        
        User32.INSTANCE.EnumWindows((hwnd, pointer) -> {
            char[] windowText = new char[512];
            User32.INSTANCE.GetWindowText(hwnd, windowText, 512);
            String windowTitle = Native.toString(windowText);

            if (windowTitle.endsWith(".xlsx") || windowTitle.endsWith(".xls")) {
                files.add(windowTitle);
            }
            return true;
        }, null);

        return files;
    }

    private List<String> findMacExcelFiles() {
        List<String> files = new ArrayList<>();
        
        try {
            String[] cmd = {
                "osascript",
                "-e", "tell application \"Microsoft Excel\"",
                "-e", "set fileNames to {}",
                "-e", "repeat with i from 1 to count of workbooks",
                "-e", "set end of fileNames to name of workbook i",
                "-e", "end repeat",
                "-e", "return fileNames",
                "-e", "end tell"
            };
            
            Process process = Runtime.getRuntime().exec(cmd);
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            String line;
            
            while ((line = reader.readLine()) != null) {
                files.add(line.trim());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        
        return files;
    }
} 