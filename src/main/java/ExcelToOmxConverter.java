import org.apache.poi.ss.usermodel.*;
import javax.swing.*;
import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

public class ExcelToOmxConverter extends JFrame {

    private JTextField filesField;
    private JTextField outputDirField;
    private JLabel statusLabel;

    private JTextField filterAIField;
    private JTextField filterAOField;
    private JTextField filterDIField;
    private JTextField filterDOField;

    private List<File> selectedFiles = new ArrayList<>();
    private File outputDir = null;

    private static class Parameter {
        final String tag;
        final String description;
        Parameter(String tag, String description) {
            this.tag = tag;
            this.description = description;
        }
    }

    public ExcelToOmxConverter() {
        setTitle("Excel → OMX (AI / AO / DI / DO) + PLC");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(820, 350);
        setLayout(new BorderLayout(10, 10));

        JPanel filePanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JButton chooseFilesButton = new JButton("Выбрать Excel файлы");
        filesField = new JTextField(40);
        filesField.setEditable(false);
        filePanel.add(chooseFilesButton);
        filePanel.add(filesField);
        add(filePanel, BorderLayout.NORTH);

        JPanel centerPanel = new JPanel();
        centerPanel.setLayout(new BoxLayout(centerPanel, BoxLayout.Y_AXIS));

        JPanel dirPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JButton chooseDirButton = new JButton("Выбрать папку для сохранения");
        outputDirField = new JTextField(40);
        outputDirField.setEditable(false);
        dirPanel.add(chooseDirButton);
        dirPanel.add(outputDirField);
        centerPanel.add(dirPanel);

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

        JPanel bottomPanel = new JPanel(new BorderLayout());
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        JButton generateButton = new JButton("Создать XML");
        buttonPanel.add(generateButton);
        statusLabel = new JLabel("Ожидание выбора файлов и папки...");
        bottomPanel.add(buttonPanel, BorderLayout.CENTER);
        bottomPanel.add(statusLabel, BorderLayout.SOUTH);
        add(bottomPanel, BorderLayout.SOUTH);

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
            for (File f : chooser.getSelectedFiles()) selectedFiles.add(f);
            filesField.setText(selectedFiles.size() <= 5 ?
                    selectedFiles.stream().map(File::getName).reduce((a,b)->a+"; "+b).orElse("") :
                    "Выбрано " + selectedFiles.size() + " файлов");
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
        if (selectedFiles.isEmpty()) { statusLabel.setText("Выберите Excel-файлы!"); return; }
        if (outputDir == null || !outputDir.exists()) { statusLabel.setText("Выберите папку!"); return; }

        int success = 0, errors = 0;
        StringBuilder errMsg = new StringBuilder();

        for (int i = 0; i < selectedFiles.size(); i++) {
            try {
                processSingleExcel(selectedFiles.get(i), outputDir, i);
                success++;
            } catch (Exception ex) {
                errors++;
                errMsg.append(selectedFiles.get(i).getName()).append(": ").append(ex.getMessage()).append("\n");
            }
        }

        try { generateMergedTypesFiles(outputDir); } catch (Exception ex) {
            errors++; errMsg.append("Объединённые Types: ").append(ex.getMessage()).append("\n");
        }

        String msg = "Готово. Успешно: " + success;
        if (errors > 0) {
            msg += ", ошибок: " + errors;
            JOptionPane.showMessageDialog(this, errMsg.toString(), "Ошибки", JOptionPane.WARNING_MESSAGE);
        }
        statusLabel.setText(msg);
    }

    private void processSingleExcel(File excelFile, File baseOutputDir, int fileIndex) throws Exception {
        String fileName = excelFile.getName();
        int dotIdx = fileName.lastIndexOf('.');
        String baseName = dotIdx > 0 ? fileName.substring(0, dotIdx) : fileName;

        String filterAI = filterAIField.getText().trim();
        String filterAO = filterAOField.getText().trim();
        String filterDI = filterDIField.getText().trim();
        String filterDO = filterDOField.getText().trim();

        List<String> aiTags = new ArrayList<>(), aiDesc = new ArrayList<>();
        List<String> aoTags = new ArrayList<>(), aoDesc = new ArrayList<>();
        List<String> diTags = new ArrayList<>(), diDesc = new ArrayList<>();
        List<String> doTags = new ArrayList<>(), doDesc = new ArrayList<>();
        List<String> plcRawEntries = new ArrayList<>();
        String sheetName = "";

        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFile))) {
            Sheet sheet = workbook.getSheetAt(0);
            sheetName = sheet.getSheetName();
            for (Row row : sheet) {
                Cell cellA = row.getCell(0);
                Cell cellB = row.getCell(1);
                if (cellA == null) continue;
                String valA = cellA.getStringCellValue().trim();
                String desc = (cellB != null) ? cellB.getStringCellValue().trim() : "";

                if (valA.startsWith("AI.ST.")) {
                    String tag = valA.substring(6);
                    if (!tag.isEmpty() && matchesFilter(tag, filterAI)) { aiTags.add(tag); aiDesc.add(desc); }
                } else if (valA.startsWith("AO.ST.")) {
                    String tag = valA.substring(6);
                    if (!tag.isEmpty() && matchesFilter(tag, filterAO)) { aoTags.add(tag); aoDesc.add(desc); }
                } else if (valA.startsWith("DI.ST.")) {
                    String tag = valA.substring(6);
                    if (!tag.isEmpty() && matchesFilter(tag, filterDI)) { diTags.add(tag); diDesc.add(desc); }
                } else if (valA.startsWith("DO.ST.")) {
                    String tag = valA.substring(6);
                    if (!tag.isEmpty() && matchesFilter(tag, filterDO)) { doTags.add(tag); doDesc.add(desc); }
                } else {
                    if (!valA.startsWith("AI.") && !valA.startsWith("AO.") &&
                            !valA.startsWith("DI.") && !valA.startsWith("DO.")) {
                        plcRawEntries.add(valA);
                    }
                }
            }
        }

        generateTypeFile(baseOutputDir, baseName, "AI", aiTags, aiDesc, "float32");
        generateTypeFile(baseOutputDir, baseName, "AO", aoTags, aoDesc, "float32");
        generateTypeFile(baseOutputDir, baseName, "DI", diTags, diDesc, "bool");
        generateTypeFile(baseOutputDir, baseName, "DO", doTags, doDesc, "bool");
        generateTypesFile(baseOutputDir, baseName, sheetName, "AI", aiTags, aiDesc, "float32");
        generateTypesFile(baseOutputDir, baseName, sheetName, "AO", aoTags, aoDesc, "float32");
        generateTypesFile(baseOutputDir, baseName, sheetName, "DI", diTags, diDesc, "bool");
        generateTypesFile(baseOutputDir, baseName, sheetName, "DO", doTags, doDesc, "bool");
        generatePlcFile(baseOutputDir, baseName, plcRawEntries, fileIndex, excelFile.getName());
    }

    private boolean matchesFilter(String tag, String filter) {
        return filter.isEmpty() || tag.contains(filter);
    }

    private void generateTypeFile(File baseDir, String baseName, String type,
                                  List<String> tags, List<String> descriptions,
                                  String dataType) throws IOException {
        if (tags.isEmpty()) return;
        File targetDir = new File(baseDir, "imports/Server/" + baseName + "/" + type);
        Files.createDirectories(targetDir.toPath());
        String xml = buildOmxXml(tags, descriptions, dataType);
        Files.writeString(new File(targetDir, baseName + "_" + type + "_ST.omx-export").toPath(), xml);
    }

    private void generateTypesFile(File baseDir, String baseName, String sheetName, String type,
                                   List<String> tags, List<String> descriptions,
                                   String dataType) throws IOException {
        if (tags.isEmpty()) return;
        File typesDir = new File(baseDir, "imports/Server/" + baseName + "/" + type + "/Types");
        Files.createDirectories(typesDir.toPath());
        String translitBase = transliterate(baseName);
        String translitUnderscore = translitBase.replace('.', '_').replaceAll("([A-Za-z])(\\d)", "$1_$2");
        StringBuilder xml = new StringBuilder();
        xml.append("<omx xmlns=\"system\" migration=\"41\" xmlns:ct=\"automation.control\" xmlns:r=\"automation.reference\">\n");
        xml.append("  <namespace name=\"").append(type).append(".").append(translitBase).append(".").append(translitUnderscore).append("\" uuid=\"\">\n");
        xml.append("    <ct:socket-type name=\"ST\" access-level=\"public\" uuid=\"\">\n");
        for (int i = 0; i < tags.size(); i++) {
            xml.append("      <ct:socket-parameter name=\"").append(escapeXml(tags.get(i)))
                    .append("\" type=\"").append(dataType).append("\" uuid=\"\">\n");
            xml.append("        <attribute type=\"unit.System.Attributes.Description\" value=\"")
                    .append(escapeXml(descriptions.get(i))).append("\" />\n");
            xml.append("      </ct:socket-parameter>\n");
        }
        xml.append("    </ct:socket-type>\n");
        String plc = type + "_PLC", serv = type + "_Serv";
        xml.append("    <ct:type name=\"").append(plc).append("\" abstract=\"false\" aspect=\"Aspects.PLC\" access-level=\"public\" uuid=\"\">\n");
        xml.append("      <ct:socket name=\"ST\" access-level=\"public\" access-scope=\"global\" direction=\"out\" type=\"").append(translitUnderscore).append(".ST\" uuid=\"\" />\n");
        xml.append("    </ct:type>\n");
        xml.append("    <ct:type name=\"").append(serv).append("\" abstract=\"false\" aspect=\"Aspects.IOS\" original=\"").append(plc).append("\" access-level=\"public\" uuid=\"\">\n");
        xml.append("      <ct:socket-bind source=\"_").append(plc).append(".ST\" target=\"ST\" />\n");
        xml.append("      <r:ref name=\"_").append(plc).append("\" type=\"").append(plc).append("\" const-access=\"false\" aspected=\"true\" uuid=\"\" />\n");
        xml.append("      <ct:socket name=\"ST\" access-level=\"public\" access-scope=\"global\" direction=\"out\" type=\"").append(translitUnderscore).append(".ST\" uuid=\"\" />\n");
        xml.append("    </ct:type>\n");
        xml.append("  </namespace>\n</omx>\n");
        Files.writeString(new File(typesDir, sheetName + ".omx-export").toPath(), xml.toString());
    }

    private void generateMergedTypesFiles(File baseOutputDir) throws IOException {
        generateMergedTypeFile("AI", filterAIField.getText().trim(), "float32", baseOutputDir);
        generateMergedTypeFile("AO", filterAOField.getText().trim(), "float32", baseOutputDir);
        generateMergedTypeFile("DI", filterDIField.getText().trim(), "bool", baseOutputDir);
        generateMergedTypeFile("DO", filterDOField.getText().trim(), "bool", baseOutputDir);
    }

    private void generateMergedTypeFile(String type, String filter, String dataType, File baseOutputDir) throws IOException {
        Map<String, Map<String, List<Parameter>>> hierarchy = new LinkedHashMap<>();
        for (File excelFile : selectedFiles) {
            String baseName = getBaseName(excelFile.getName());
            String translitFull = transliterate(baseName).replace('.', '_').replaceAll("([A-Za-z])(\\d)", "$1_$2");
            int underscoreIdx = translitFull.indexOf('_');
            String upperNS = (underscoreIdx > 0) ? translitFull.substring(0, underscoreIdx) : translitFull;
            List<Parameter> params = new ArrayList<>();
            try (Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFile))) {
                for (Row row : workbook.getSheetAt(0)) {
                    Cell cellA = row.getCell(0);
                    Cell cellB = row.getCell(1);
                    if (cellA == null) continue;
                    String valA = cellA.getStringCellValue().trim();
                    if (!valA.startsWith(type + ".ST.")) continue;
                    String tag = valA.substring((type + ".ST.").length());
                    if (tag.isEmpty() || !matchesFilter(tag, filter)) continue;
                    String desc = (cellB != null) ? cellB.getStringCellValue().trim() : "";
                    params.add(new Parameter(tag, desc));
                }
            } catch (Exception ignored) {}
            if (!params.isEmpty()) {
                hierarchy.computeIfAbsent(upperNS, k -> new LinkedHashMap<>()).put(translitFull, params);
            }
        }
        if (hierarchy.isEmpty()) return;
        StringBuilder xml = new StringBuilder();
        xml.append("<omx xmlns=\"system\" migration=\"41\" xmlns:ct=\"automation.control\" xmlns:r=\"automation.reference\">\n");
        xml.append("  <namespace name=\"").append(type).append("\" uuid=\"\">\n");
        for (Map.Entry<String, Map<String, List<Parameter>>> upperEntry : hierarchy.entrySet()) {
            xml.append("    <namespace name=\"").append(upperEntry.getKey()).append("\" uuid=\"\">\n");
            for (Map.Entry<String, List<Parameter>> nsEntry : upperEntry.getValue().entrySet()) {
                String fullNS = nsEntry.getKey();
                xml.append("      <namespace name=\"").append(fullNS).append("\" uuid=\"\">\n");
                xml.append("        <ct:socket-type name=\"ST\" access-level=\"public\" uuid=\"\">\n");
                for (Parameter p : nsEntry.getValue()) {
                    xml.append("          <ct:socket-parameter name=\"").append(escapeXml(p.tag))
                            .append("\" type=\"").append(dataType).append("\" uuid=\"\">\n");
                    xml.append("            <attribute type=\"unit.System.Attributes.Description\" value=\"")
                            .append(escapeXml(p.description)).append("\" />\n");
                    xml.append("          </ct:socket-parameter>\n");
                }
                xml.append("        </ct:socket-type>\n");
                String plc = type + "_PLC", serv = type + "_Serv";
                xml.append("        <ct:type name=\"").append(plc).append("\" abstract=\"false\" aspect=\"Aspects.PLC\" access-level=\"public\" uuid=\"\">\n");
                xml.append("          <ct:socket name=\"ST\" access-level=\"public\" access-scope=\"global\" direction=\"out\" type=\"").append(fullNS).append(".ST\" uuid=\"\" />\n");
                xml.append("        </ct:type>\n");
                xml.append("        <ct:type name=\"").append(serv).append("\" abstract=\"false\" aspect=\"Aspects.IOS\" original=\"").append(plc).append("\" access-level=\"public\" uuid=\"\">\n");
                xml.append("          <ct:socket-bind source=\"_").append(plc).append(".ST\" target=\"ST\" />\n");
                xml.append("          <r:ref name=\"_").append(plc).append("\" type=\"").append(plc).append("\" const-access=\"false\" aspected=\"true\" uuid=\"\" />\n");
                xml.append("          <ct:socket name=\"ST\" access-level=\"public\" access-scope=\"global\" direction=\"out\" type=\"").append(fullNS).append(".ST\" uuid=\"\" />\n");
                xml.append("        </ct:type>\n");
                xml.append("      </namespace>\n");
            }
            xml.append("    </namespace>\n");
        }
        xml.append("  </namespace>\n</omx>\n");
        File serverDir = new File(baseOutputDir, "imports/Server");
        Files.createDirectories(serverDir.toPath());
        Files.writeString(new File(serverDir, "Types_" + type + ".omx-export").toPath(), xml.toString());
    }

    private void generatePlcFile(File baseOutputDir, String baseName,
                                 List<String> rawEntries, int fileIndex, String sourceFileName) throws IOException {
        if (rawEntries.isEmpty()) return;
        Map<String, Set<String>> rootToSubs = new LinkedHashMap<>();
        for (String entry : rawEntries) {
            String[] parts = entry.split("\\.");
            if (parts.length < 2) continue;
            String root = parts[0], sub = parts[1];
            if (root.isEmpty() || sub.isEmpty()) continue;
            rootToSubs.computeIfAbsent(root, k -> new LinkedHashSet<>()).add(sub);
        }
        if (rootToSubs.isEmpty()) return;

        String translitName = transliterate(baseName).replace('.', '_').replaceAll("([A-Za-z])(\\d)", "$1_$2");
        String ip = "192.168.1." + (fileIndex + 1);
        char firstChar = translitName.charAt(0);
        String buildingNumber = Character.isDigit(firstChar) ? String.valueOf(firstChar) : "1";

        StringBuilder xml = new StringBuilder();
        xml.append("<omx xmlns=\"system\" migration=\"41\" xmlns:dp=\"automation.deployment\"")
                .append(" xmlns:eth=\"automation.ethernet\" xmlns:modbus=\"automation.modbus\"")
                .append(" xmlns:ct=\"automation.control\">\n");
        xml.append("  <dp:computer name=\"").append(translitName).append("\" uuid=\"\">\n");
        xml.append("    <eth:ethernet-adapter name=\"EthernetAdapter\" address=\"").append(ip)
                .append("\" use-dhcp=\"false\" network=\"Ethernet\" uuid=\"\" />\n");
        xml.append("    <dp:external-runtime name=\"Runtime\" uuid=\"\">\n");
        xml.append("      <dp:application-object name=\"Application\" access-level=\"public\" uuid=\"\">\n");
        xml.append("        <modbus:modbus-link-map name=\"ModbusAddressMap\" file=\"Building_")
                .append(buildingNumber).append("\\").append(translitName).append(".xml\" uuid=\"\">\n");
        xml.append("          <attribute type=\"unit.System.Attributes.LinkMapSettings\"")
                .append(" value=\"{" +
                        "&quot;LinkArea&quot;:0," +
                        "&quot;CommandsSegment&quot;:1," +
                        "&quot;MeasuringSegment&quot;:2," +
                        "&quot;OptimizeCategory&quot;:false," +
                        "&quot;RegulationSegment&quot;:3," +
                        "&quot;SeparateObjects&quot;:true," +
                        "&quot;SignalingSegment&quot;:0," +
                        "&quot;StartAddresses&quot;:{}}\" />\n");
        xml.append("        </modbus:modbus-link-map>\n");

        for (Map.Entry<String, Set<String>> entry : rootToSubs.entrySet()) {
            String root = entry.getKey();
            Set<String> subs = entry.getValue();
            List<String> validSubs = new ArrayList<>();
            for (String sub : subs) {
                String baseType = getBaseType(root, sub);
                if (baseType == null) {
                    System.out.println("Неизвестный подобъект: " + sub + " для корня " + root +
                            " в файле " + sourceFileName);
                } else {
                    validSubs.add(sub);
                }
            }
            if (validSubs.isEmpty()) continue;

            xml.append("        <ct:object name=\"").append(root)
                    .append("\" access-scope=\"global\" access-level=\"public\" uuid=\"\">\n");
            for (String sub : validSubs) {
                String baseType = getBaseType(root, sub);
                xml.append("          <ct:object name=\"").append(sub)
                        .append("\" base-type=\"").append(baseType)
                        .append("\" access-scope=\"global\" aspect=\"Aspects.PLC\" access-level=\"public\" uuid=\"\" />\n");
            }
            xml.append("        </ct:object>\n");
        }

        xml.append("      </dp:application-object>\n");
        xml.append("      <modbus:modbus-tcp-slave name=\"ModbusTcpSlave\" number=\"255\"")
                .append(" switch-bytes=\"false\" switch-float-words=\"false\" switch-int-words=\"false\"")
                .append(" switch-str-bytes=\"false\" port=\"502\" address-map=\"Application.ModbusAddressMap\"")
                .append(" text-encoding=\"ibm-5347_P100-1998\" uuid=\"\" />\n");
        xml.append("    </dp:external-runtime>\n");
        xml.append("  </dp:computer>\n");
        xml.append("</omx>\n");

        File plcDir = new File(baseOutputDir, "imports/Server/" + baseName);
        Files.createDirectories(plcDir.toPath());
        Files.writeString(new File(plcDir, baseName + "_PLC.omx-export").toPath(), xml.toString());
    }

    private String getBaseType(String root, String sub) {
        // Точные совпадения для часто встречающихся
        switch (sub) {
            case "Common": return "Types.Common.Common_PLC";
            case "Scheduler":
            case "DailyShedule": return "Types.Ventilation.DailyShedule.DailyShedule_PLC";
            case "WaterHeater": return "Types.Ventilation.WaterHeater.WaterHeater_PLC";
            case "ElectricHeater": return "Types.Ventilation.ElectricHeater.ElectricHeater_PLC";
            case "Recup": return "Types.Ventilation.RecupPlate.RecupPlate_PLC";
            case "CCU": return "Types.Ventilation.CCU.CCU_PLC";
            case "Leak": return "Types.Leak.Leak_PLC";
            case "Light": return "Types.Light.Light_PLC";
            case "MBSensor": return "Types.MBSensor.MBSensor_PLC";
            case "Core":
                if (root.startsWith("PV")) return "Types.Ventilation.CoreSupExh.CoreSupExh_PLC";
                else if (root.startsWith("V")) return "Types.Ventilation.CoreExh.CoreExh_PLC";
                else return "Types.Ventilation.CoreSup.CoreSup_PLC";
        }

        // Фильтры и вентиляторы с возможными числовыми суффиксами
        if (sub.startsWith("FilterSup") || sub.startsWith("FilterExh")) {
            return "Types.Ventilation.Filter.Filter_PLC";
        }
        if (sub.startsWith("FanSup") || sub.startsWith("FanExh")) {
            return "Types.Ventilation.FanSingle.FanSingle_PLC";
        }
        // Простые вентиляторы (KB...)
        if (sub.startsWith("KB")) {
            return "Types.Ventilation.FanSimple.FanSimple_PLC";
        }

        // Если ничего не подошло — возвращаем null (неизвестный тип)
        return null;
    }

    private String buildOmxXml(List<String> tags, List<String> descriptions, String dataType) {
        StringBuilder sb = new StringBuilder();
        sb.append("<omx xmlns=\"system\" migration=\"41\" xmlns:ct=\"automation.control\">\n");
        sb.append("  <ct:socket-type name=\"ST\" access-level=\"public\" uuid=\"\">\n");
        for (int i = 0; i < tags.size(); i++) {
            sb.append("    <ct:socket-parameter name=\"").append(escapeXml(tags.get(i)))
                    .append("\" type=\"").append(dataType).append("\" uuid=\"\">\n");
            sb.append("      <attribute type=\"unit.System.Attributes.Description\" value=\"")
                    .append(escapeXml(descriptions.get(i))).append("\" />\n");
            sb.append("    </ct:socket-parameter>\n");
        }
        sb.append("  </ct:socket-type>\n</omx>\n");
        return sb.toString();
    }

    private String transliterate(String cyrillic) {
        String[][] map = {{"Щ","SH"},{"щ","sh"},{"А","A"},{"а","a"},{"Б","B"},{"б","b"},
                {"В","V"},{"в","v"},{"Г","G"},{"г","g"},{"Д","D"},{"д","d"},{"Е","E"},{"е","e"},
                {"Ё","Yo"},{"ё","yo"},{"Ж","Zh"},{"ж","zh"},{"З","Z"},{"з","z"},{"И","I"},{"и","i"},
                {"Й","Y"},{"й","y"},{"К","K"},{"к","k"},{"Л","L"},{"л","l"},{"М","M"},{"м","m"},
                {"Н","N"},{"н","n"},{"О","O"},{"о","o"},{"П","P"},{"п","p"},{"Р","R"},{"р","r"},
                {"С","S"},{"с","s"},{"Т","T"},{"т","t"},{"У","U"},{"у","u"},{"Ф","F"},{"ф","f"},
                {"Х","Kh"},{"х","kh"},{"Ц","Ts"},{"ц","ts"},{"Ч","Ch"},{"ч","ch"},{"Ш","Sh"},{"ш","sh"},
                {"Ъ",""},{"ъ",""},{"Ы","Y"},{"ы","y"},{"Ь",""},{"ь",""},{"Э","E"},{"э","e"},
                {"Ю","Yu"},{"ю","yu"},{"Я","Ya"},{"я","ya"}};
        String res = cyrillic;
        for (String[] pair : map) res = res.replace(pair[0], pair[1]);
        return res;
    }

    private String escapeXml(String s) {
        if (s == null) return "";
        return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                .replace("\"","&quot;").replace("'","&apos;");
    }

    private String getBaseName(String fileName) {
        int dot = fileName.lastIndexOf('.');
        return dot > 0 ? fileName.substring(0, dot) : fileName;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(ExcelToOmxConverter::new);
    }
}