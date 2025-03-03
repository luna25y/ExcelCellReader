package com.excelreader;

import com.sun.speech.freetts.Voice;
import com.sun.speech.freetts.VoiceManager;
import org.apache.poi.ss.usermodel.*;

public class ExcelReader {
    private Voice voice;
    private boolean isReading;

    public ExcelReader() {
        System.setProperty("freetts.voices", "com.sun.speech.freetts.en.us.cmu_us_kal.KevinVoiceDirectory");
        voice = VoiceManager.getInstance().getVoice("kevin16");
        if (voice != null) {
            voice.allocate();
        }
    }

    public void startReading(String column) {
        isReading = true;
        // 实现读取Excel文件并朗读指定列的逻辑
    }

    public void stopReading() {
        isReading = false;
        if (voice != null) {
            voice.deallocate();
        }
    }
} 