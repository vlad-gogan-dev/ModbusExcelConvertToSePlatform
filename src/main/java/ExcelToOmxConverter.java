import org.apache.poi.ss.usermodel.*;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class ExcelToOmxConverter extends JFrame {

    private JTextField filesField;
    private JTextField outputDirField;
    private JLabel statusLabel;

    // Поля фильтров
    private JTextField filterAIField;
    private JTextField filterAOField;
    private JTextField filterDIField;
    private JTextField filterDOField;

    private List<File> selectedFiles = new ArrayList<>();
    private File outputDir = null;

    public ExcelToOmxConverter() {
        setTitle("Excel → OMX (AI / AO / DI / DO)");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(820, 350);
        setLayout(new BorderLayout(10, 10));

        // Верх: выбор файлов
        JPanel filePanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JButton chooseFilesButton = new JButton("Выбрать Excel файлы");
        filesField = new JTextField(40);
        filesField.setEditable(false);
        filePanel.add(chooseFilesButton);
        filePanel.add(filesField);
        add(filePanel, BorderLayout.NORTH);

        // Центральная часть: папка + фильтры
        JPanel centerPanel = new JPanel();
        centerPanel.setLayout(new BoxLayout(centerPanel, BoxLayout.Y_AXIS));

        // Панель выбора папки
        JPanel dirPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JButton chooseDirButton = new JButton("Выбрать папку для сохранения");
        outputDirField = new JTextField(40);
        outputDirField.setEditable(false);
        dirPanel.add(chooseDirButton);
        dirPanel.add(outputDirField);
        centerPanel.add(dirPanel);

        // Панель фильтров (4 поля)
        JPanel filterPanel = new JPanel(new GridLayout(2, 4, 5, 5));
        filterPanel.setBorder(BorderFactory.createTitledBorder("Генерация по соответствию (фильтры)"));
        filterPanel.add(new JLabel("Фильтр AI:", SwingConstants.CENTER));
        filterPanel.add(new JLabel("Фильтр AO:", SwingConstants.CENTER));
        filterPanel.add(new JLabel("Фильтр DI:", SwingConstants.CENTER));
        filterPanel.add(new JLabel("Фильтр DO:", SwingConstants.CENTER));

        filterAIField = new JTextField();
        filterAOField = new JTextField();
        filterDIField = new JTextField();
        filterDOField = new JTextField();

        filterPanel.add(filterAIField);
        filterPanel.add(filterAOField);
        filterPanel.add(filterDIField);
        filterPanel.add(filterDOField);
        centerPanel.add(filterPanel);

        add(centerPanel, BorderLayout.CENTER);

        // Низ: кнопка генерации и статус
        JPanel bottomPanel = new JPanel(new BorderLayout());
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        JButton generateButton = new JButton("Создать XML");
        buttonPanel.add(generateButton);
        statusLabel = new JLabel("Ожидание выбора файлов и папки...");
        bottomPanel.add(buttonPanel, BorderLayout.CENTER);
        bottomPanel.add(statusLabel, BorderLayout.SOUTH);
        add(bottomPanel, BorderLayout.SOUTH);

        // Обработчики
        chooseFilesButton.addActionListener(this::chooseFiles);
        chooseDirButton.addActionListener(this::chooseOutputDir);
        generateButton.addActionListener(this::generateXml);

        setLocationRelativeTo(null);
        setVisible(true);
    }

    private void chooseFiles(ActionEvent e) {
        JFileChooser chooser = new JFileChooser();
        chooser.setMultiSelectionEnabled(true);
        chooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
                "Excel files", "xlsx", "xls"));
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            selectedFiles.clear();
            for (File f : chooser.getSelectedFiles()) {
                selectedFiles.add(f);
            }
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
            msg += ", ошибок: " + errorCount;
            JOptionPane.showMessageDialog(this, errors.toString(),
                    "Ошибки при обработке", JOptionPane.WARNING_MESSAGE);
        }
        statusLabel.setText(msg);
    }

    private void processSingleExcel(File excelFile, File baseOutputDir) throws Exception {
        String fileName = excelFile.getName();
        int dotIdx = fileName.lastIndexOf('.');
        String baseName = dotIdx > 0 ? fileName.substring(0, dotIdx) : fileName;

        String filterAI = filterAIField.getText().trim();
        String filterAO = filterAOField.getText().trim();
        String filterDI = filterDIField.getText().trim();
        String filterDO = filterDOField.getText().trim();

        List<String> aiTags = new ArrayList<>();
        List<String> aiDesc = new ArrayList<>();
        List<String> aoTags = new ArrayList<>();
        List<String> aoDesc = new ArrayList<>();
        List<String> diTags = new ArrayList<>();
        List<String> diDesc = new ArrayList<>();
        List<String> doTags = new ArrayList<>();
        List<String> doDesc = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFile))) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                Cell cellA = row.getCell(0);
                Cell cellB = row.getCell(1);
                if (cellA == null) continue;
                String valueA = cellA.getStringCellValue().trim();
                String desc = (cellB != null) ? cellB.getStringCellValue().trim() : "";

                if (valueA.startsWith("AI.ST.")) {
                    String tag = valueA.substring(6);
                    if (!tag.isEmpty() && matchesFilter(tag, filterAI)) {
                        aiTags.add(tag);
                        aiDesc.add(desc);
                    }
                } else if (valueA.startsWith("AO.ST.")) {
                    String tag = valueA.substring(6);
                    if (!tag.isEmpty() && matchesFilter(tag, filterAO)) {
                        aoTags.add(tag);
                        aoDesc.add(desc);
                    }
                } else if (valueA.startsWith("DI.ST.")) {
                    String tag = valueA.substring(6);
                    if (!tag.isEmpty() && matchesFilter(tag, filterDI)) {
                        diTags.add(tag);
                        diDesc.add(desc);
                    }
                } else if (valueA.startsWith("DO.ST.")) {
                    String tag = valueA.substring(6);
                    if (!tag.isEmpty() && matchesFilter(tag, filterDO)) {
                        doTags.add(tag);
                        doDesc.add(desc);
                    }
                }
            }
        }

        generateTypeFile(baseOutputDir, baseName, "AI", aiTags, aiDesc, "float32");
        generateTypeFile(baseOutputDir, baseName, "AO", aoTags, aoDesc, "float32");
        generateTypeFile(baseOutputDir, baseName, "DI", diTags, diDesc, "bool");
        generateTypeFile(baseOutputDir, baseName, "DO", doTags, doDesc, "bool");
    }

    private boolean matchesFilter(String tag, String filter) {
        if (filter.isEmpty()) return true;
        return tag.contains(filter);
    }

    private void generateTypeFile(File baseDir, String baseName, String type,
                                  List<String> tags, List<String> descriptions,
                                  String dataType) throws IOException {
        if (tags.isEmpty()) return;

        File targetDir = new File(baseDir, "imports/Server/" + baseName + "/" + type);
        Files.createDirectories(targetDir.toPath());

        String xmlContent = buildOmxXml(tags, descriptions, dataType);
        String outputFileName = baseName + "_" + type + "_ST.omx-export";
        File outputFile = new File(targetDir, outputFileName);
        Files.writeString(outputFile.toPath(), xmlContent);
    }

    private String buildOmxXml(List<String> tags, List<String> descriptions, String dataType) {
        StringBuilder sb = new StringBuilder();
        sb.append("<omx xmlns=\"system\" migration=\"41\" xmlns:ct=\"automation.control\">\n");
        sb.append("  <ct:socket-type name=\"ST\" access-level=\"public\" uuid=\"\">\n");

        for (int i = 0; i < tags.size(); i++) {
            sb.append("    <ct:socket-parameter name=\"")
                    .append(escapeXml(tags.get(i)))
                    .append("\" type=\"").append(dataType).append("\" uuid=\"\">\n");
            sb.append("      <attribute type=\"unit.System.Attributes.Description\" value=\"")
                    .append(escapeXml(descriptions.get(i)))
                    .append("\" />\n");
            sb.append("    </ct:socket-parameter>\n");
        }

        sb.append("  </ct:socket-type>\n");
        sb.append("</omx>\n");
        return sb.toString();
    }

    private String escapeXml(String s) {
        if (s == null) return "";
        return s
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