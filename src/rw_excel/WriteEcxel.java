package rw_excel;

// *****  IMPORTACIONES ****
import java.io.File;
import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.IndexedColors; 
import org.apache.poi.ss.util.CellRangeAddress;

public class WriteEcxel {
    //*** Estilos posibles ***
    //Se puede asignar algunos a una celda enviados como un arreglo de string al momento de crear la celda
//   {    "decimal", "fecha",
//        "negrita", "cursiva", "centrar", "alinear_derecha",
//        "t_12", "t_14", 
//        "t_blanco", "t_rojo", "t_verde",
//        "cafe", "rojo", "verde", "azul", "gris25%", "gris40%", "gris80%", 
//        "bordes" }
    
    public XSSFWorkbook libro = new XSSFWorkbook();
    public XSSFSheet hoja;
    public String urlCreacion="";       //Es la direccion donde se guradara el libro INCLUIDO el nombre del archivo con su extension respectiva 
    ArrayList<Integer> filas = new  ArrayList<Integer>();       //lista de filas que contiene el libro, sirve para la creacion de celdas
    ArrayList<Integer> columnas = new  ArrayList<Integer>();    //Lista de columnas que posee el libro, solo sirve para auto- ajustar
    
    //para conversion de cadenas a fechas
    SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
    
    WriteEcxel(String url){
        this.urlCreacion=url;
        this.hoja=this.libro.createSheet("Hoja1");
    }
    
    private CellStyle getEstilo(String[] estilosIN){
        CellStyle  estilo = this.libro.createCellStyle();
        Font font = this.libro.createFont();
        for(String item: estilosIN){           
            switch (item) 
            {
                case "negrita":  
                    font.setBold(true);
                    break;
                case "cursiva":  
                    font.setItalic(true);
                    break;
                case "t_12":  
                    font.setFontHeightInPoints((short)12);
                    break;
                case "t_14":  
                    font.setFontHeightInPoints((short)14);
                    break;
                case "t_blanco":  
                    font.setColor(IndexedColors.WHITE.getIndex());
                    break;
                case "t_rojo":  
                    font.setColor(IndexedColors.RED.getIndex());
                    break;
                case "t_verde":  
                    font.setColor(IndexedColors.GREEN.getIndex());
                    break;
                case "cafe":
                    estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    estilo.setFillForegroundColor(IndexedColors.BROWN.getIndex());
                    break;
                case "rojo":  
                    estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    estilo.setFillForegroundColor(IndexedColors.RED.getIndex());
                    break;
                case "verde":  
                    estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    estilo.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
                    break;
                case "azul":  
                    estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    estilo.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                    break;
                case "gris25%":  
                    estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    estilo.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                    break;
                case "gris40%":  
                    estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    estilo.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
                    break;
                case "gris80%":  
                    estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    estilo.setFillForegroundColor(IndexedColors.GREY_80_PERCENT.getIndex());
                    break;
                case "bordes":  
                    estilo.setBorderBottom(BorderStyle.THIN);
                    estilo.setBorderLeft(BorderStyle.THIN);
                    estilo.setBorderRight(BorderStyle.THIN);
                    estilo.setBorderTop(BorderStyle.THIN);
                    break;
                case "centrar":  
                    estilo.setAlignment(HorizontalAlignment.CENTER);
                    break;
                case "alinear_derecha":  
                    estilo.setAlignment(HorizontalAlignment.RIGHT);
                    break;
                case "decimal":  
                    estilo.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                    break;
                case "fecha":
                    CreationHelper createHelper = this.libro.getCreationHelper();
                    estilo.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));
                    break;    
                default: 
                    //dayString = "Dia inválido";
                    break;
            }
        }
        estilo.setFont(font);
        return estilo;
    }
    
    public void crearCelda(int fila, int columna , String valor, String[] estilo) throws ParseException{
        //El arreglo de estilo es un arreglo de cadenas que seran similares a la variable global estilos
        XSSFRow row;   
        XSSFCell cell;
        //Si la fila no esta en la lista entonces la crea (y toma la referencia), si ya existe toma la referencia
        if (!this.filas.contains(fila)) {
            row = this.hoja.createRow(fila);
            this.filas.add(fila);
        }else{
            row = this.hoja.getRow(fila);
        }
        cell = row.createCell(columna);
        cell.setCellStyle(this.getEstilo(estilo)); //aplica estillo a la celda
        
        if (Arrays.asList(estilo).contains("decimal")) {
            cell.setCellValue(Double.parseDouble(valor));
        }else{
            if (Arrays.asList(estilo).contains("fecha")) {
                try{
                    Date date = this.formatter.parse(valor); 
                    cell.setCellValue(date);
                }catch (ParseException e) {
                    e.printStackTrace();
                }
            }else{
                cell.setCellValue(valor);
            }
        }     
        if (!this.columnas.contains(columna)) {
            this.columnas.add(columna);
        }
    }
    
    public void guardarLibro(){
        //********* Auto Ajustar texto a lo ancho
        for (int i = 0; i < this.columnas.size(); i++) {
            this.hoja.autoSizeColumn(this.columnas.get(i));
        }
        //********* Creacion de Archivo 
        File excelFile;
        excelFile = new File(this.urlCreacion); // Referenciando a la ruta y el archivo Excel a crear
        try (FileOutputStream fileOuS = new FileOutputStream(excelFile)) {
            if (excelFile.exists()) { // Si el archivo existe lo eliminaremos
                excelFile.delete();
                System.out.println("Archivo eliminado.!");
            }
            //coloca todo el contenido que hay en nuestro libro actual en el archivo creado
            this.libro.write(fileOuS);
            fileOuS.flush();
            fileOuS.close();
            System.out.println("Archivo Creado.!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    public void combinarCentrar(int fila, int columnaInicial, int columnaFinal){
        this.hoja.addMergedRegion(new CellRangeAddress(fila,fila,columnaInicial,columnaFinal));
    }
}
