import mx.com.sura.facturacion.commons.bean.ReporteFacturacion;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;

/**
 * @author Marcos Santiago Leonardo
 * Meltsan Solutions
 * Description: crea reporte de comprobantes en excel
 * Date: 2/6/19
 */
public final class ExcelWritter {
    public static final String[] columns = {
            "Oficina", "Ramo", "Póliza", "Fecha Emisión Póliza",
            "Tipo comprobante", "UUID", "CFDI Relacionado", "Fecha Emisión Comprobante",
            "Forma Pago", "Método Pago", "RFC", "Razón Social",
            "Prima Neta", "Derechos", "Recargos", "IVA",
            "Total"
    };

    public static void createReport(String pathFile, List<ReporteFacturacion> listComprobante) throws Exception {

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
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }


        /*formato de fecha*/
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd/MM/yyyy"));

        /*Creando contenido*/
        int rowNum = 1;
        for (ReporteFacturacion comprobante : listComprobante) {
            Row row = sheet.createRow(rowNum++);
            int colum = 0;

            if (comprobante.getDsTipoComprobante().equalsIgnoreCase("INGRESO")) {
                for (Field field : comprobante.getClass().getDeclaredFields()) {
                    field.setAccessible(true);
                    if (StringUtils.isBlank(String.valueOf(field.get(comprobante)))
                            || String.valueOf(field.get(comprobante)).equalsIgnoreCase("null")) {
                        row.createCell(colum++).setCellValue("");
                    } else {
                        if (field.getName().equalsIgnoreCase("cfdiRelacionado")) {
                            row.createCell(colum++).setCellValue("");
                        } else {
                            row.createCell(colum++).setCellValue(String.valueOf(field.get(comprobante)));
                        }
                    }


                }
            } else if (comprobante.getDsTipoComprobante().equalsIgnoreCase("COMPLEMENTO")) {
                createRow(comprobante, row, colum);
            } else if (comprobante.getDsTipoComprobante().equalsIgnoreCase("EGRESO")) {
                createRow(comprobante, row, colum);
            }
        }

        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        try {
            FileOutputStream fileOutputStream = new FileOutputStream(pathFile);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            throw new Exception(e.getMessage());
        }
    }

    private static void createRow(ReporteFacturacion comprobante, Row row, int colum) throws IllegalAccessException {
        for (Field field : comprobante.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            if (StringUtils.isBlank(String.valueOf(field.get(comprobante)))
                    || String.valueOf(field.get(comprobante)).equalsIgnoreCase("null")) {
                row.createCell(colum++).setCellValue("");
            } else {
                row.createCell(colum++).setCellValue(String.valueOf(field.get(comprobante)));
            }

        }
    }
}
