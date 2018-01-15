import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.filechooser.FileFilter;
import javax.swing.plaf.ColorUIResource;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Arrays;

public class ExcelParserPanel extends JPanel
        implements ActionListener {


    private JButton openButton, parseButton;
    private JLabel filePathFielLabel, comboBoxLabel;
    private JComboBox<String> headerComboBox;
    private JTextArea openPathTextArea;
    private JFileChooser openFileChooser;
    private File sourceFile;
    private ExcelParser parser;
    private String filePath, selectedColumn;

    private static final String PIWIK_HEADER = "Referring Action Name";
    private static final String INVALID_EXCEL_WARNING = "Bitte wählen Sie eine *.xlsx-Datei als Quelle aus";
    private static final String NO_COLUMN_SELECTED_WARNING = "Bitte wählen Sie eine Spalte aus, für die eine Chart generiert werden soll";


    private ExcelParserPanel() {
        super(new GridBagLayout());

        filePathFielLabel = new JLabel("Wähle die Excel-Datei aus: ", SwingConstants.LEFT);

        GridBagConstraints c = new GridBagConstraints();
        c.fill = GridBagConstraints.HORIZONTAL;
        c.insets = new Insets(40, 40, 0, 40);
        c.gridx = 0;
        c.gridy = 0;

        add(filePathFielLabel, c);

        openPathTextArea = new JTextArea(1, 50);
        openPathTextArea.setMargin(new Insets(5, 5, 5, 5));
        openPathTextArea.setEditable(false);

        Border border = BorderFactory.createLineBorder(Color.BLACK);
        openPathTextArea.setBorder(BorderFactory.createCompoundBorder(border,
                BorderFactory.createEmptyBorder(5, 5, 5, 5)));

        c.insets = new Insets(10, 40, 0, 40);
        c.fill = GridBagConstraints.NONE;
        c.gridx = 0;
        c.gridy = 1;

        add(openPathTextArea, c);

        openButton = new JButton("Datei auswählen...");
        openButton.addActionListener(this);

        c.insets = new Insets(10, 0, 0, 40);
        c.gridx = 3;
        c.gridy = 1;
        c.gridwidth = 1;

        add(openButton, c);

        comboBoxLabel = new JLabel("Wähle die Spalte aus, die ausgewertet werden soll: ", SwingConstants.LEFT);

        c.fill = GridBagConstraints.HORIZONTAL;
        c.insets = new Insets(30, 40, 0, 40);
        c.gridx = 0;
        c.gridy = 2;

        add(comboBoxLabel, c);

        UIManager.put("ComboBox.background", new ColorUIResource(Color.WHITE));

        String[] array = {""};

        headerComboBox = new JComboBox(array);
        headerComboBox.setSelectedIndex(0);
        headerComboBox.addActionListener(this);
        headerComboBox.setEnabled(false);

        c.insets = new Insets(10, 40, 50, 40);
        c.gridx = 0;
        c.gridy = 3;
        c.gridwidth = 3;

        add(headerComboBox, c);


        parseButton = new JButton("Chart generieren");
        parseButton.addActionListener(this);
        parseButton.setEnabled(false);

        c.insets = new Insets(10, 0, 50, 40);
        c.gridx = 3;
        c.gridy = 3;

        add(parseButton, c);

        initFileChooser();
    }

    private void initFileChooser() {
        openFileChooser = new JFileChooser();
        openFileChooser.addChoosableFileFilter(new FileFilter() {

            public String getDescription() {
                return "Microsoft Excel Documents (*.xlsx)";
            }

            public boolean accept(File f) {
                if (f.isDirectory()) {
                    return true;
                } else {
                    return f.getName().toLowerCase().endsWith(".xlsx");
                }
            }
        });
    }

    public void actionPerformed(ActionEvent e) {

        if (e.getSource() == openButton) {

            int returnVal = openFileChooser.showOpenDialog(ExcelParserPanel.this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {

                File selectedSourceFile = openFileChooser.getSelectedFile();

                if (selectedSourceFile != null) {
                    if (checkForFileExtension(selectedSourceFile.getPath(), ".xlsx")) {
                        sourceFile = selectedSourceFile;

                        filePath = sourceFile.getPath();
                        openPathTextArea.setText("");
                        openPathTextArea.append(filePath);
                        openPathTextArea.setCaretPosition(openPathTextArea.getDocument().getLength());

                        //init parser
                        parser = new ExcelParser(filePath);

                        //check for "piwik" columns to determine wether a correct excel file was selected
                        if (checkForPiwikColumns(parser.getHeaderRowAsArray())) {

                            //fill ComboBox with headers
                            DefaultComboBoxModel model = new DefaultComboBoxModel(parser.getHeaderRowAsArray());
                            headerComboBox.setModel(model);
                            parseButton.setEnabled(true);

                            headerComboBox.setEnabled(isSelectedPathsValid());
                        } else {
                            showComboBoxWarningDialog(INVALID_EXCEL_WARNING);
                            openPathTextArea.setText("");
                        }
                    } else {
                        showComboBoxWarningDialog(INVALID_EXCEL_WARNING);
                    }
                }
            }
        } else if (e.getSource() == headerComboBox) {

            selectedColumn = headerComboBox.getSelectedItem().toString();

        } else if (e.getSource() == parseButton) {

            try {
                if (!selectedColumn.isEmpty() && selectedColumn != null) {
                    parser.createChartForSelectedColumn(selectedColumn);
                } else {
                    showComboBoxWarningDialog(NO_COLUMN_SELECTED_WARNING);
                }
            } catch (NullPointerException exc) {
                showComboBoxWarningDialog(NO_COLUMN_SELECTED_WARNING);
            }
        }
    }


    //TODO Adjust this method accordingly in order to check, whether a Piwik Excel file was selected
    private boolean checkForPiwikColumns(String[] headers) {

        boolean isPiwikExcel = false;

        for (String header : headers) {
            if (Arrays.asList(headers).contains(PIWIK_HEADER)) {
                isPiwikExcel = true;
            }
        }
        return isPiwikExcel;
    }

    private void showComboBoxWarningDialog(String msg) {
        JOptionPane.showMessageDialog(ExcelParserPanel.this,
                msg,
                "Warnung",
                JOptionPane.WARNING_MESSAGE);
    }

    private boolean isSelectedPathsValid() {

        String excelFileName = openPathTextArea.getText();

        if (checkForFileExtension(excelFileName, ".xlsx")) {
            return true;
        }
        return false;
    }

    private boolean checkForFileExtension(String filename, String extension) {

        if (extension.equals(".xlsx")) {

            if (!filename.isEmpty()) {
                if (filename.lastIndexOf(".") != -1 && filename.lastIndexOf(".") != 0) {

                    String currentExtension = filename.substring(filename.lastIndexOf("."));

                    if (currentExtension.equals(extension)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }


    private static void createAndShowGUI() {

        JFrame frame = new JFrame("Wähle Excel-Datei aus");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        frame.add(new ExcelParserPanel());

        frame.pack();
        frame.setVisible(true);
    }

    public static void main(String[] args) {

        SwingUtilities.invokeLater(() -> {
            UIManager.put("swing.boldMetal", Boolean.FALSE);
            createAndShowGUI();
        });
    }

}