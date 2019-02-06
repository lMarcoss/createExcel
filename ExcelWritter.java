package mx.com;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author Marcos Santiago Leonardo
 * Meltsan Solutions
 * Description: escribe en archivo excel
 * Date: 2/6/19
 */
public class ExcelWritter {
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();

        CreationHelper creationHelper = workbook.getCreationHelper();

        Sheet sheet = workbook.createSheet();

        /* estilos para cabecera */
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.BLACK.getIndex());


        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        Row headerRow = sheet.createRow(0);

        //create a cells
        for (int i = 0; i < 16; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue("HOLA");
            cell.setCellStyle(headerCellStyle);
        }


        /*formato de fecha*/
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd/MM/yyyy"));

        /*Creando contenido*/
        int rowNum = 1;
        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(rowNum++);
            for (int j = 0; j < 16; j++) {
                row.createCell(j).setCellValue(j);
            }
        }

        for (int i = 0; i < 16; i++) {
            sheet.autoSizeColumn(i);
        }

        try {
            FileOutputStream fileOutputStream = new FileOutputStream("/Users/lMarcoss/workspace-sura/projects/FacturacionServices/data.xls");
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
