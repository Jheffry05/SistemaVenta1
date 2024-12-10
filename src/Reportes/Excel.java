package Reportes;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import javax.swing.JOptionPane;
import Conexion.Conexion;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

    public static void reporte() {
        Workbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Productos");

        // Estilo de título
        CellStyle tituloEstilo = book.createCellStyle();
        tituloEstilo.setAlignment(HorizontalAlignment.CENTER);
        Font fuenteTitulo = book.createFont();
        fuenteTitulo.setFontName("Arial");
        fuenteTitulo.setBold(true);
        fuenteTitulo.setFontHeightInPoints((short) 14);
        tituloEstilo.setFont(fuenteTitulo);

        // Fila de título
        Row filaTitulo = sheet.createRow(0);
        Cell celdaTitulo = filaTitulo.createCell(0);
        celdaTitulo.setCellStyle(tituloEstilo);
        celdaTitulo.setCellValue("Reporte de Productos");

        // Fila de encabezado
        String[] cabecera = new String[]{"Código", "Nombre", "Precio", "Existencia"};
        CellStyle headerStyle = book.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);

        Font font = book.createFont();
        font.setFontName("Arial");
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setFontHeightInPoints((short) 12);
        headerStyle.setFont(font);

        Row filaEncabezados = sheet.createRow(1);
        for (int i = 0; i < cabecera.length; i++) {
            Cell celdaEncabezado = filaEncabezados.createCell(i);
            celdaEncabezado.setCellStyle(headerStyle);
            celdaEncabezado.setCellValue(cabecera[i]);
        }

        // Conexión a la base de datos y generación del reporte
        try (Connection conn = new Conexion().getConnection();
             PreparedStatement ps = conn.prepareStatement("SELECT codigo, nombre, precio, stock FROM productos");
             ResultSet rs = ps.executeQuery()) {

            int numFilaDatos = 2;
            CellStyle datosEstilo = book.createCellStyle();
            datosEstilo.setBorderBottom(BorderStyle.THIN);
            datosEstilo.setBorderLeft(BorderStyle.THIN);
            datosEstilo.setBorderRight(BorderStyle.THIN);
            datosEstilo.setBorderTop(BorderStyle.THIN);

            while (rs.next()) {
                Row filaDatos = sheet.createRow(numFilaDatos);
                for (int i = 0; i < cabecera.length; i++) {
                    Cell celdaDatos = filaDatos.createCell(i);
                    celdaDatos.setCellStyle(datosEstilo);
                    celdaDatos.setCellValue(rs.getString(i + 1)); // Las columnas SQL empiezan en 1
                }
                numFilaDatos++;
            }

            // Autoajustar tamaño de columnas
            for (int i = 0; i < cabecera.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Guardar archivo en la carpeta Descargas
            String home = System.getProperty("user.home");
            String fileName = "productos.xlsx";
            File file = new File(home + "/Downloads/" + fileName);

            // Verificar si el archivo ya existe
            if (file.exists()) {
                JOptionPane.showMessageDialog(null, "El archivo ya existe. Será reemplazado.");
            }

            try (FileOutputStream fileOut = new FileOutputStream(file)) {
                book.write(fileOut);
            }

            // Intentar abrir el archivo generado
            if (Desktop.isDesktopSupported()) {
                Desktop.getDesktop().open(file);
            } else {
                JOptionPane.showMessageDialog(null, "El archivo se ha guardado en Descargas, pero no se puede abrir automáticamente.");
            }

            JOptionPane.showMessageDialog(null, "Reporte generado MiniMarket DETODITO.");
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Error al consultar los datos: " + ex.getMessage());
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Error al escribir el archivo: " + ex.getMessage());
        }
    }
}
