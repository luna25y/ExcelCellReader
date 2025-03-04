package com.excelreader;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import javafx.application.Platform;
import javafx.scene.layout.Priority;

public class ExcelReaderApp extends Application {
    private ExcelReader excelReader;
    private ExcelFileWatcher fileWatcher;
    private ListView<String> excelFilesList;
    private boolean isReading = false;

    @Override
    public void start(Stage primaryStage) {
        excelReader = new ExcelReader();
        fileWatcher = new ExcelFileWatcher();

        VBox root = new VBox(10);
        root.setPadding(new Insets(10));

        // 控制按钮
        HBox controlBox = new HBox(10);
        Button playButton = new Button("播放");
        Button stopButton = new Button("停止");
        TextField columnInput = new TextField();
        columnInput.setPromptText("输入列号 (例如: A)");
        columnInput.setPrefWidth(150);

        controlBox.getChildren().addAll(playButton, stopButton, columnInput);

        // Excel文件列表
        excelFilesList = new ListView<>();
        VBox.setVgrow(excelFilesList, Priority.ALWAYS);

        // 链接图标
        HBox linksBox = new HBox(10);
        ImageView githubIcon = new ImageView(new Image(getClass().getResourceAsStream("/icons/github.png")));
        ImageView websiteIcon = new ImageView(new Image(getClass().getResourceAsStream("/icons/website.png")));
        
        githubIcon.setFitHeight(24);
        githubIcon.setFitWidth(24);
        websiteIcon.setFitHeight(24);
        websiteIcon.setFitWidth(24);

        githubIcon.setOnMouseClicked(e -> getHostServices().showDocument("https://github.com/luna25y"));
        websiteIcon.setOnMouseClicked(e -> getHostServices().showDocument("https://www.luna25y.com"));

        linksBox.getChildren().addAll(githubIcon, websiteIcon);

        // 退出按钮
        Button exitButton = new Button("退出");
        exitButton.setOnAction(e -> Platform.exit());

        root.getChildren().addAll(controlBox, excelFilesList, linksBox, exitButton);

        // 事件处理
        playButton.setOnAction(e -> {
            String column = columnInput.getText().trim().toUpperCase();
            if (!column.isEmpty()) {
                isReading = true;
                excelReader.startReading(column);
            }
        });

        stopButton.setOnAction(e -> {
            isReading = false;
            excelReader.stopReading();
        });

        // 监听Excel文件变化
        fileWatcher.setOnExcelFilesChanged(files -> {
            Platform.runLater(() -> {
                excelFilesList.getItems().clear();
                excelFilesList.getItems().addAll(files);
            });
        });
        fileWatcher.start();

        Scene scene = new Scene(root, 400, 600);
        scene.getStylesheets().add(getClass().getResource("/styles.css").toExternalForm());

        primaryStage.setTitle("Excel Reader");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    @Override
    public void stop() {
        fileWatcher.stop();
        excelReader.stopReading();
        Platform.exit();
    }
} 