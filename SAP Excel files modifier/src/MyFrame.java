import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Iterator;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MyFrame extends JFrame implements ActionListener {
    private final JButton button1;
    private final JButton button2;
    private final JButton button3;
    private final JButton exitButton;
    List<File> selectedFilesList = new ArrayList<>();

    public MyFrame() {
        setTitle("ONE RECOUVREMENT");
        setIconImage(new ImageIcon("logo.png").getImage());
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());
        setResizable(false);

        DefaultListModel<String> listModel = new DefaultListModel<>();
        JList<String> filesList = new JList<>(listModel);
        filesList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);

        JPanel panel1 = new JPanel();
        panel1.setLayout(new FlowLayout(FlowLayout.CENTER, 10, 20));

        button1 = createStyledButton("Select files", "Select files", "file_icon.png");
        button2 = createStyledButton("Modify files", "Edit files", "edit_icon.png");
        button3 = createStyledButton("Create a global file", "Create global file", "document_icon.png");
        exitButton = createStyledButton();

        button1.addActionListener(this);
        button2.addActionListener(this);
        button3.addActionListener(this);
        exitButton.addActionListener(this);

        panel1.add(button1);
        panel1.add(button2);
        panel1.add(button3);

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 10, 20));
        buttonPanel.add(exitButton);

        add(panel1, BorderLayout.CENTER);
        add(buttonPanel, BorderLayout.SOUTH);

        JLabel titleLabel = new JLabel("Welcome to ONE RECOUVREMENT");
        titleLabel.setFont(new Font("Verdana", Font.ITALIC, 28));
        titleLabel.setForeground(new Color(0, 0, 0));
        titleLabel.setHorizontalAlignment(SwingConstants.CENTER);
        add(titleLabel, BorderLayout.NORTH);

        pack();

        setMinimumSize(new Dimension(450, getHeight()));

        setLocationRelativeTo(null);

        setVisible(true);
    }

    private JButton createStyledButton() {
        JButton button = new JButton("Exit");
        button.setBackground(new Color(231, 76, 60));
        button.setForeground(Color.white);
        button.setFocusPainted(false);
        button.setBorder(BorderFactory.createEmptyBorder(12, 24, 12, 24));
        button.setFont(new Font("Arial", Font.BOLD, 14));
        button.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        button.setToolTipText("Exit");

        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                button.setBackground(new Color(231, 52, 26));
            }

            public void mouseExited(java.awt.event.MouseEvent evt) {
                button.setBackground(new Color(231, 76, 60));
            }
        });

        FontMetrics fontMetrics = button.getFontMetrics(button.getFont());
        int textWidth = fontMetrics.stringWidth("Exit");
        button.setPreferredSize(new Dimension(textWidth + 60, button.getPreferredSize().height));

        return button;
    }

    private JButton createStyledButton(String text, String tooltip, String iconFileName) {
        JButton button = new JButton(text);
        button.setBackground(new Color(52, 152, 219));
        button.setForeground(Color.white);
        button.setFocusPainted(false);
        button.setBorder(BorderFactory.createEmptyBorder(12, 24, 12, 24));
        button.setFont(new Font("Arial", Font.BOLD, 14));
        button.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        button.setToolTipText(tooltip);

        try {
            ImageIcon icon = new ImageIcon(iconFileName);
            Image scaledIcon = icon.getImage().getScaledInstance(18, 18, Image.SCALE_SMOOTH);
            button.setIcon(new ImageIcon(scaledIcon));
        } catch (Exception e) {
            e.printStackTrace();
        }

        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                button.setBackground(new Color(41, 128, 185));
            }

            public void mouseExited(java.awt.event.MouseEvent evt) {
                button.setBackground(new Color(52, 152, 219));
            }
        });

        FontMetrics fontMetrics = button.getFontMetrics(button.getFont());
        int textWidth = fontMetrics.stringWidth(text);
        button.setPreferredSize(new Dimension(textWidth + 60, button.getPreferredSize().height));

        return button;
    }

    List<String> inputFiles = new ArrayList<>();

    @Override
    public void actionPerformed(ActionEvent e) {
        if (e.getSource() == button1) {
            selectFiles();
        } else if (e.getSource() == button2) {
            modifyFiles();
        } else if (e.getSource() == button3) {
            mergeExcelFiles(inputFiles);
        } else if (e.getSource() == exitButton) {
            System.exit(0);
        }
    }

    public void selectFiles() {
        JFileChooser file = new JFileChooser();
        selectedFilesList.clear();
        file.setDialogTitle("Select 8 Excel Files");
        file.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
        file.setMultiSelectionEnabled(true);
        file.setFileSelectionMode(JFileChooser.FILES_ONLY); // Ensure only files, not directories, can be selected

        int response = file.showOpenDialog(null);
        if (response == JFileChooser.APPROVE_OPTION) {
            selectedFilesList.clear();
            File[] selectedFiles = file.getSelectedFiles();

            if (selectedFiles.length == 8) {
                Collections.addAll(selectedFilesList, selectedFiles);
                ImageIcon successIcon = new ImageIcon("successIcon.png");
                JOptionPane.showMessageDialog(this, "The files are successfully selected.", "Information", JOptionPane.PLAIN_MESSAGE, successIcon);
            } else {
                JOptionPane.showMessageDialog(this, "Please select exactly 8 files.", "Erreur", JOptionPane.ERROR_MESSAGE);
            }
        } else {
            JOptionPane.showMessageDialog(this, "The user cancels file selection", "Information", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private void modifyFiles() {
        inputFiles.clear();
        JFileChooser file = new JFileChooser();
        file.setDialogTitle("Choose the destination folder to save the modified files:");
        file.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int userSelection = file.showSaveDialog(null);
        if (userSelection != JFileChooser.APPROVE_OPTION) {
            JOptionPane.showMessageDialog(this, "The user has cancelled the file modification process.", "Information", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        File destinationFolder = file.getSelectedFile();
        String outputFolderPath = destinationFolder.getAbsolutePath();
        if (!outputFolderPath.endsWith(File.separator)) {
            outputFolderPath += File.separator;
        }
        for (File inputFile : selectedFilesList) {
            String outputFilepath;
            try (FileInputStream fis = new FileInputStream(inputFile);
                 Workbook workbook = WorkbookFactory.create(fis)) {
                sheet = workbook.getSheetAt(0);
                int lastColumnIndex = getLastColumnIndex(sheet);
                headerRow = sheet.getRow(0);
                if (headerRow == null) {
                    headerRow = sheet.createRow(0);
                }
                Cell cell1 = headerRow.createCell(lastColumnIndex);
                cell1.setCellValue("Année");
                Cell cell2 = headerRow.createCell(lastColumnIndex + 1);
                cell2.setCellValue("Nature");
                Cell cell3 = headerRow.createCell(lastColumnIndex + 2);
                cell3.setCellValue("Concaténation");
                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);

                    Cell yearCell = row.createCell(lastColumnIndex);
                    Cell natureCell = row.createCell(lastColumnIndex + 1);

                    String fileName = inputFile.getName().toLowerCase();
                    if (fileName.contains("tr")) {
                        natureCell.setCellValue("Travaux");
                    } else {
                        natureCell.setCellValue("Energie");
                    }
                    Pattern pattern = Pattern.compile("20\\d{2}");
                    Matcher matcher = pattern.matcher(fileName);

                    if (matcher.find()) {
                        String year = matcher.group();
                        yearCell.setCellValue(year);
                    } else {
                        JOptionPane.showMessageDialog(this, "Make sure all selected files contain a year in their name.", "Information", JOptionPane.ERROR_MESSAGE);
                    }
                }
                for (Cell cell : headerRow) {
                    if (cell.getStringCellValue().equalsIgnoreCase("Montant échu")) {
                        mEchuColumnIndex = cell.getColumnIndex();
                        mRegleColumnIndex = mEchuColumnIndex + 1;
                        break;
                    }
                }
                if (mEchuColumnIndex != -1) {
                    NvMontant(sheet, mEchuColumnIndex);
                    NvMontant(sheet, mRegleColumnIndex);
                }
                int ClCmptColumnIndex = -1;
                for (Cell cell : headerRow) {
                    if (cell.getStringCellValue().equalsIgnoreCase("Classe de compte")) {
                        ClCmptColumnIndex = cell.getColumnIndex();
                        break;
                    }
                }
                String fileName = inputFile.getName().toLowerCase();
                if (fileName.contains("tr")) {
                    ClCmptTr(sheet, ClCmptColumnIndex);
                } else if (fileName.contains("en")) {
                    ClCmptEn(sheet, ClCmptColumnIndex);
                }
                int agenceColumnIndex = -1;
                for (Cell cell : headerRow) {
                    if (cell.getStringCellValue().equalsIgnoreCase("GpeStrReg")) {
                        agenceColumnIndex = cell.getColumnIndex();
                        break;
                    }
                }
                if (agenceColumnIndex != -1) {
                    nvAgence(sheet, agenceColumnIndex);
                }
                int TypeClientColumnIndex = -1;
                for (Cell cell : headerRow) {
                    if (cell.getStringCellValue().equalsIgnoreCase("Type client")) {
                        TypeClientColumnIndex = cell.getColumnIndex();
                        break;
                    }
                }
                if (TypeClientColumnIndex != -1) {
                    nvTypeClient(sheet, TypeClientColumnIndex);
                }
                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    Cell concatCell = row.createCell(lastColumnIndex + 2);
                    concatCell.setCellValue(row.getCell(1).getStringCellValue() + row.getCell(2).getStringCellValue() + row.getCell(3).getStringCellValue() + row.getCell(8).getStringCellValue());
                }
                CellStyle commonStyle = workbook.createCellStyle();
                commonStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                commonStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                commonStyle.setBorderTop(BorderStyle.THIN);
                commonStyle.setBorderBottom(BorderStyle.THIN);
                commonStyle.setBorderRight(BorderStyle.THIN);
                commonStyle.setBorderLeft(BorderStyle.THIN);
                commonStyle.setAlignment(HorizontalAlignment.CENTER);
                Row row0 = sheet.getRow(0);
                for (Cell cell : row0) {
                    cell.setCellStyle(commonStyle);
                    int indedx = cell.getColumnIndex();
                    applyCenterAlignment(sheet, indedx);
                }
                int numColumns = sheet.getRow(0).getLastCellNum();
                int[] maxColumnWidths = new int[numColumns];

                for (Row row : sheet) {
                    for (int columnIndex = 0; columnIndex < numColumns; columnIndex++) {
                        Cell cell = row.getCell(columnIndex);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue();
                            int length = cellValue.length();
                            if (length > maxColumnWidths[columnIndex]) {
                                maxColumnWidths[columnIndex] = length;
                            }
                        }
                    }
                }
                for (int columnIndex = 0; columnIndex < numColumns; columnIndex++) {
                    sheet.setColumnWidth(columnIndex, (maxColumnWidths[columnIndex] + 2) * 256);
                }
                deleteRows(sheet);

                String outputFileName = "Modified_" + inputFile.getName();
                outputFilepath = outputFolderPath + outputFileName;
                try (FileOutputStream outputStream = new FileOutputStream(outputFolderPath + outputFileName)) {
                    workbook.write(outputStream);
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
            inputFiles.add(outputFilepath);
        }
        ImageIcon successIcon = new ImageIcon("successIcon.png");
        JOptionPane.showMessageDialog(this, "The selected files are successfully modified", "Information", JOptionPane.PLAIN_MESSAGE, successIcon);
    }

    int mEchuColumnIndex = -1, mRegleColumnIndex = -1;
    Sheet sheet;
    Row headerRow;

    private static int getLastColumnIndex(Sheet sheet) {
        int lastColumnIndex = 0;
        for (Row row : sheet) {
            int lastCellIndex = row.getLastCellNum();
            if (lastCellIndex > lastColumnIndex) {
                lastColumnIndex = lastCellIndex;
            }
        }
        return lastColumnIndex;
    }

    private static void applyCenterAlignment(Sheet sheet, int columnIndex) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(columnIndex);
            if (cell != null) {
                cell.setCellStyle(cellStyle);
            }
        }
    }

    private static void NvMontant(Sheet sheet, int columnIndex) {
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(columnIndex);
            cell.setCellValue(cell.getNumericCellValue() / 1000);
        }
    }

    static List<Integer> allRowsToDelete = new ArrayList<>();

    private static void ClCmptEn(Sheet sheet, int columnIndex) {
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(columnIndex);
            String val = cell.getStringCellValue();
            switch (val) {
                case "Autres Etablissement Publics" -> cell.setCellValue("Stés nationales");
                case "Les agents ONE" -> cell.setCellValue("Particuliers");
                case "Multi-Contrats (Régl Reg) Autres" -> cell.setCellValue("Multi-Contrats (Régl Regional)");
                case "" -> allRowsToDelete.add(rowIndex);
            }
        }
    }

    private static void ClCmptTr(Sheet sheet, int columnIndex) {
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(columnIndex);
            String val = cell.getStringCellValue();
            switch (val) {
                case "Autres Etablissement Publics" -> cell.setCellValue("Stés nationales");
                case "Clients occasionnels" -> cell.setCellValue("Particuliers");
                case "PALAIS ROYAL" -> cell.setCellValue("Administrations");
                case "Multi-Contrats (Régl Reg) Autres" -> cell.setCellValue("Multi-Contrats (Régl Regional)");
                case "" -> allRowsToDelete.add(rowIndex);
            }
        }
    }

    private static void nvAgence(Sheet sheet, int columnIndex) {
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(columnIndex);
            String val = cell.getStringCellValue();
            switch (val) {
                case "AGENCE DE SERVICES PROVINCIALE LAAYOUNE" -> cell.setCellValue("Agence de Services Provinciale Laâyoune");
                case "AGENCE DE SERVICES LAKHSSASS" -> cell.setCellValue("AGENCE DE SERVICES T. LAKHSSASS");
                case "SUCCURSALE BIR GANDOUZ" -> cell.setCellValue("Succursale Bir Gandouz");
            }
        }
    }

    private static void nvTypeClient(Sheet sheet, int columnIndex) {
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(columnIndex);
            String val = cell.getStringCellValue();
            switch (val) {
                case "CB", "CX", "EB", "EC", "EP", "NA", "PP" -> cell.setCellValue("BT");
                case "CM", "EM", "GC", "HT" -> cell.setCellValue("MT");
                case "" -> allRowsToDelete.add(rowIndex);
            }
        }
    }

    private static void deleteRows(Sheet sheet) {
        List<Integer> rowsToDelete = allRowsToDelete.stream().distinct().toList();

        for (Integer rowIndex : rowsToDelete) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                row.setZeroHeight(true);
            }
        }
    }

    private static void deleteRows(Sheet sheet, int index) {
        Row row = sheet.getRow(index);
        if (row != null) {
            row.setZeroHeight(true);
        }
    }

    private void mergeExcelFiles(List<String> filePaths) {
        JFileChooser save = new JFileChooser();
        save.setDialogTitle("Choose the destination folder to save the global file:");
        save.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        int userSelection = save.showSaveDialog(null);
        if (userSelection != JFileChooser.APPROVE_OPTION) {
            JOptionPane.showMessageDialog(this,
                    "The user has cancelled the creation of the global file.",
                    "Information", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        File destinationFolder = save.getSelectedFile();
        outputFolderPath = destinationFolder.getAbsolutePath();
        if (!outputFolderPath.endsWith(File.separator)) {
            outputFolderPath += File.separator;
        }

        try (Workbook outputWorkbook = new XSSFWorkbook()) {
            Iterator<Cell> headerIterator = null;
            Sheet inputSheet;
            for (String filePath : filePaths) {
                try (FileInputStream fis = new FileInputStream(filePath)) {
                    Workbook inputWorkbook = new XSSFWorkbook(fis);
                    for (int i = 0; i < inputWorkbook.getNumberOfSheets(); i++) {
                        inputSheet = inputWorkbook.getSheetAt(i);
                        Sheet outputSheet = outputWorkbook.getSheet(inputSheet.getSheetName());
                        if (outputSheet == null) {
                            outputSheet = outputWorkbook.createSheet(inputSheet.getSheetName());
                        }
                        int startRow = (headerIterator != null) ? 1 : 0;

                        for (Row inputRow : inputSheet) {
                            if (startRow == 0) {
                                headerIterator = inputRow.cellIterator();
                                Row outputHeaderRow = outputSheet.createRow(0);
                                while (headerIterator.hasNext()) {
                                    Cell inputHeaderCell = headerIterator.next();
                                    Cell outputHeaderCell = outputHeaderRow.createCell
                                            (inputHeaderCell.getColumnIndex(), inputHeaderCell.getCellType());
                                    outputHeaderCell.setCellValue(inputHeaderCell.getStringCellValue());
                                }
                            } else {
                                Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                                for (Cell inputCell : inputRow) {
                                    Cell outputCell = outputRow.createCell(inputCell.getColumnIndex(), inputCell.getCellType());
                                    switch (inputCell.getCellType()) {
                                        case STRING -> outputCell.setCellValue(inputCell.getStringCellValue());
                                        case NUMERIC -> outputCell.setCellValue(inputCell.getNumericCellValue());
                                        case BOOLEAN -> outputCell.setCellValue(inputCell.getBooleanCellValue());
                                        case FORMULA -> outputCell.setCellFormula(inputCell.getCellFormula());
                                        default -> {
                                        }
                                    }
                                }
                            }
                            startRow++;
                        }
                    }
                }
            }

            Sheet sheet = outputWorkbook.getSheetAt(0);
            CellStyle commonStyle = outputWorkbook.createCellStyle();
            commonStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            commonStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            commonStyle.setBorderTop(BorderStyle.THIN);
            commonStyle.setBorderBottom(BorderStyle.THIN);
            commonStyle.setBorderRight(BorderStyle.THIN);
            commonStyle.setBorderLeft(BorderStyle.THIN);
            commonStyle.setAlignment(HorizontalAlignment.CENTER);
            Row row0 = sheet.getRow(0);
            for (Cell cell : row0) {
                cell.setCellStyle(commonStyle);
                int indedx = cell.getColumnIndex();
                applyCenterAlignment(sheet, indedx);
            }

            int numColumns = sheet.getRow(0).getLastCellNum();
            int[] maxColumnWidths = new int[numColumns];
            for (Row row : sheet) {
                for (int columnIndex = 0; columnIndex < numColumns; columnIndex++) {
                    Cell cell = row.getCell(columnIndex);
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        int length = cellValue.length();
                        if (length > maxColumnWidths[columnIndex]) {
                            maxColumnWidths[columnIndex] = length;
                        }
                    }
                }
            }
            for (int columnIndex = 0; columnIndex < numColumns; columnIndex++) {
                sheet.setColumnWidth(columnIndex, (maxColumnWidths[columnIndex] + 2) * 256);
            }

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().isEmpty()) {
                        deleteRows(sheet, cell.getRowIndex());
                    }
                }
            }
            outputFileName = "Global.xlsx";
            try (FileOutputStream fos = new FileOutputStream(outputFolderPath + outputFileName)) {
                outputWorkbook.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        ImageIcon successIcon = new ImageIcon("successIcon.png");
        JOptionPane.showMessageDialog(this,
                "The global file has been created successfully",
                "Information", JOptionPane.PLAIN_MESSAGE, successIcon);
    }

    String outputFileName;
    String outputFolderPath;
}
