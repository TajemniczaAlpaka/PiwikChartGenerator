import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class ExcelParser {

    private XSSFSheet sheet;
    private Row headerRow;
    private String path;


    public ExcelParser(String filePath) {

        if (!filePath.isEmpty()) {
            this.path = filePath;
            initExcelSheet();
        }
    }

    public String[] getHeaderRowAsArray(){

        String[] headers = new String[headerRow.getPhysicalNumberOfCells() + 1];

        Iterator it = headerRow.iterator();

        headers[0] = "";
        int i = 1;

        while(it.hasNext()) {
            headers[i] = ((Cell) it.next()).getStringCellValue();
            i++;
        }
        return headers;
    }

    private void initExcelSheet() {
        try {
            FileInputStream excelFile = new FileInputStream(new File(path));
            Workbook workbook = new XSSFWorkbook(excelFile);
            sheet = (XSSFSheet) workbook.getSheetAt(0);
            headerRow = sheet.getRow(0);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     *
     * Finds the row for the selected column name
     *
     * @param cellContent The selected column name
     * @return The values in the column
     */
    private ArrayList<String> findRow(String cellContent) {

        if (sheet != null && headerRow != null) {
            ArrayList<String> columnValues = new ArrayList<>();
            for (Cell headerRowCell : headerRow) {
                if ((headerRowCell.getCellTypeEnum() == CellType.STRING) &&
                        headerRowCell.getRichStringCellValue().getString().trim().equals(cellContent)) {
                    int columnIndex = headerRowCell.getColumnIndex();
                    for (int i = 1; i <= sheet.getPhysicalNumberOfRows(); i++) {
                        Row row = sheet.getRow(i);
                        if (row != null) {
                            Cell cell = row.getCell(columnIndex);
                            if (cell != null && (cell.getCellTypeEnum() == CellType.STRING)) {
                                columnValues.add(cell.getStringCellValue());
                            }
                        }
                    }
                }
            }
            return columnValues;
        }
        return null;
    }

    public void setSelectedColumn(String value){

    }

    public void createChartForSelectedColumn(String selectedColumn){
        if(!selectedColumn.isEmpty()){
            ArrayList<String> columnValues = findRow(selectedColumn);
            PieChart demo = new PieChart("Statistik", selectedColumn, columnValues);
            demo.pack();
            demo.setVisible(true);
        }

    }
}
