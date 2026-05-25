import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class ExcelToOmxConverter extends JFrame {

    private JTextField filesField;          // показывает количество или имена выбранных файлов
    private JTextField outputDirField;      // путь к папке для сохранения
    private JLabel statusLabel;

    private List<File> selectedFiles = new ArrayList<>();
    private File outputDir = null;

    public ExcelToOmxConverter() {
        setTitle("Excel → OMX (множественный выбор)");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(750, 200);
        setLayout(new BorderLayout(10, 10));

        // --- Верхняя панель: выбор файлов ---
        JPanel filePanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JButton chooseFilesButton = new JButton("Выбрать Excel файлы");
        filesField = new JTextField(40);
        filesField.setEditable(false);
        filePanel.add(chooseFilesButton);
        filePanel.add(filesField);
        add(filePanel, BorderLayout.NORTH);

        // --- Средняя панель: выбор папки ---
        JPanel dirPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JButton chooseDirButton = new JButton("Выбрать папку для сохранения");
        outputDirField = new JTextField(40);
        outputDirField.setEditable(false);
        dirPanel.add(chooseDirButton);
        dirPanel.add(outputDirField);
        add(dirPanel, BorderLayout.CENTER);

        // --- Нижняя панель: кнопка генерации + статус ---
        JPanel bottomPanel = new JPanel(new BorderLayout());
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        JButton generateButton = new JButton("Создать XML");
        buttonPanel.add(generateButton);
        statusLabel = new JLabel("Ожидание выбора файлов и папки...");
        bottomPanel.add(buttonPanel, BorderLayout.CENTER);
        bottomPanel.add(statusLabel, BorderLayout.SOUTH);
        add(bottomPanel, BorderLayout.SOUTH);

        // --- Обработчики событий ---
        chooseFilesButton.addActionListener(this::chooseFiles);
        chooseDirButton.addActionListener(this::chooseOutputDir);
        generateButton.addActionListener(this::generateXml);

        setLocationRelativeTo(null);
        setVisible(true);
    }

    // Выбор нескольких Excel-файлов
    private void chooseFiles(ActionEvent e) {
        JFileChooser chooser = new JFileChooser();
        chooser.setMultiSelectionEnabled(true);
        chooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
                "Excel files", "xlsx", "xls"));
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File[] files = chooser.getSelectedFiles();
            selectedFiles.clear();
            for (File f : files) {
                selectedFiles.add(f);
            }
            // Показываем список имён (или количество)
            if (selectedFiles.size() <= 5) {
                StringBuilder sb = new StringBuilder();
                for (File f : selectedFiles) {
                    sb.append(f.getName()).append("; ");
                }
                filesField.setText(sb.toString());
            } else {
                filesField.setText("Выбрано " + selectedFiles.size() + " файлов");
            }
            statusLabel.setText("Выбрано файлов: " + selectedFiles.size());
        }
    }

    // Выбор папки назначения
    private void chooseOutputDir(ActionEvent e) {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setAcceptAllFileFilterUsed(false);
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            outputDir = chooser.getSelectedFile();
            outputDirField.setText(outputDir.getAbsolutePath());
            statusLabel.setText("Папка выбрана: " + outputDir.getName());
        }
    }

    // Генерация XML для всех выбранных файлов
    private void generateXml(ActionEvent e) {
        if (selectedFiles.isEmpty()) {
            statusLabel.setText("Сначала выберите хотя бы один Excel-файл!");
            return;
        }
        if (outputDir == null || !outputDir.exists()) {
            statusLabel.setText("Сначала выберите существующую папку для сохранения!");
            return;
        }

        int successCount = 0;
        int errorCount = 0;
        StringBuilder errors = new StringBuilder();

        for (File excelFile : selectedFiles) {
            try {
                processSingleExcel(excelFile, outputDir);
                successCount++;
            } catch (Exception ex) {
                errorCount++;
                errors.append(excelFile.getName()).append(": ").append(ex.getMessage()).append("\n");
            }
        }

        String msg = "Готово. Успешно: " + successCount;
        if (errorCount > 0) {
            msg += ", ошибок: " + errorCount + " (см. ниже)";
            errors.insert(0, msg + "\n\n");
            JOptionPane.showMessageDialog(this, errors.toString(),
                    "Результат обработки", JOptionPane.WARNING_MESSAGE);
        } else {
            msg += ". Файлы сохранены в " + outputDir.getAbsolutePath();
        }
        statusLabel.setText(msg);
    }

    // Обработка одного Excel-файла и сохранение XML
    private void processSingleExcel(File excelFile, File outputFolder) throws Exception {
        // Извлекаем базовое имя без расширения
        String fileName = excelFile.getName();
        String baseName;
        int dotIndex = fileName.lastIndexOf('.');
        if (dotIndex > 0) {
            baseName = fileName.substring(0, dotIndex);
        } else {
            baseName = fileName;
        }

        // Имя выходного файла
        String outputFileName = baseName + "_AI_ST.omx-export";
        File outputFile = new File(outputFolder, outputFileName);

        // Строим XML
        String header = """
                <omx xmlns="system" migration="41" xmlns:ct="automation.control">
                  <ct:socket-type name="ST" access-level="public" uuid="">
                """;
        String footer = """
                  </ct:socket-type>
                </omx>
                """;

        StringBuilder parametersBlock = new StringBuilder();

        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFile))) {
            Sheet sheet = workbook.getSheetAt(0); // первый лист
            for (Row row : sheet) {
                Cell cellA = row.getCell(0); // столбец A
                Cell cellB = row.getCell(1); // столбец B

                if (cellA == null) continue;
                String valueA = cellA.getStringCellValue().trim();
                if (!valueA.startsWith("AI.ST.")) continue;

                // Извлекаем Tag (всё после "AI.ST.")
                String tag = valueA.substring(6); // длина "AI.ST." = 6
                if (tag.isEmpty()) continue;

                // Значение из столбца B (NameTag)
                String nameTag = "";
                if (cellB != null) {
                    nameTag = cellB.getStringCellValue().trim();
                }

                // Формируем узел параметра
                parametersBlock.append("    <ct:socket-parameter name=\"")
                        .append(escapeXml(tag))
                        .append("\" type=\"float32\" uuid=\"\">\n")
                        .append("      <attribute type=\"unit.System.Attributes.Description\" value=\"")
                        .append(escapeXml(nameTag))
                        .append("\" />\n")
                        .append("    </ct:socket-parameter>\n");
            }
        }

        String fullXml = header + parametersBlock.toString() + footer;
        Files.writeString(outputFile.toPath(), fullXml);
    }

    // Экранирование спецсимволов XML
    private String escapeXml(String input) {
        if (input == null) return "";
        return input
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(ExcelToOmxConverter::new);
    }
}